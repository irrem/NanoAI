using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using NanoAI.Gemini;
using System.Windows.Forms;

namespace NanoAI.Core.NLU.Handlers
{
    /// <summary>
    /// AI-powered handler for natural language commands that aren't explicitly defined
    /// </summary>
    public class SmartCommandHandler : ICommandHandler
    {
        private readonly GeminiService _geminiService;
        private readonly ICommandProcessingService _parentService;
        private readonly HttpClient _httpClient;
        
        // Custom prompt for AI-powered command processing
        private const string AI_COMMAND_PROMPT = @"
You're an AI assistant integrated with Windows that can control applications and perform tasks. I need you to:

1. Analyze the command: ""{0}""
2. Determine what the user wants to do
3. Identify which Windows app or system function is needed
4. Provide the exact command that should be executed

Reply with a JSON structure containing:
- analysis: Brief explanation of what the user wants
- appRequired: Name of the Windows app needed (or 'system' for built-in functions)
- actionType: One of [launch, search, ui_interact, file_operation, system_control]
- specificAction: What exact action to take
- parameters: Any values needed to complete the action

Example for ""send an email to John"":
{{
  ""analysis"": ""User wants to send an email to a contact named John"",
  ""appRequired"": ""outlook"",
  ""actionType"": ""launch"",
  ""specificAction"": ""compose_email"",
  ""parameters"": {{
    ""recipient"": ""John"",
    ""subject"": """",
    ""body"": """"
  }}
}}

Focus on:
- Identifying the correct app needed (Word, Excel, PowerPoint, Notepad, Calculator, etc.)
- Understanding the specific action (open, save, edit, calculate, etc.)
- Extracting relevant parameters (file names, text content, values)

Return ONLY this JSON structure with no additional explanation.
";

        public string CommandType => "smart";

        public SmartCommandHandler(GeminiService geminiService, ICommandProcessingService parentService)
        {
            _geminiService = geminiService ?? throw new ArgumentNullException(nameof(geminiService));
            _parentService = parentService ?? throw new ArgumentNullException(nameof(parentService));
            _httpClient = new HttpClient();
        }

        public bool CanHandle(GeminiCommand command)
        {
            // This handler can process any "custom" or "smart" command
            return command.CommandType.Equals("custom", StringComparison.OrdinalIgnoreCase) ||
                   command.CommandType.Equals("smart", StringComparison.OrdinalIgnoreCase);
        }

        public async Task<CommandResult> ExecuteAsync(GeminiCommand command)
        {
            // Get the raw command text
            string commandText = command.Target;
            
            if (string.IsNullOrEmpty(commandText))
            {
                return new CommandResult
                {
                    Success = false,
                    Message = "No command provided",
                    Suggestions = new List<string> { "Please provide a command to execute" }
                };
            }
            
            try
            {
                // Step 1: Process with AI to understand the command
                var smartCommand = await ProcessWithAI(commandText);
                
                // Step 2: Convert AI response to executable command
                var executableCommand = await ConvertToExecutableCommand(smartCommand, commandText);
                
                if (executableCommand != null)
                {
                    // Step 3: Execute the command using the parent service
                    return await _parentService.ExecuteCommandDirectlyAsync(executableCommand);
                }
                
                // Step 4: If we couldn't convert to a known command type, try web search
                if (smartCommand.Parameters.TryGetValue("analysis", out var analysis))
                {
                    // This seems to be a knowledge-based query, perform web search
                    if (IsQuestionOrQuery(commandText))
                    {
                        return await PerformWebSearch(commandText);
                    }
                }
                
                // Step 5: Last resort - try to launch or find an application
                string appRequired = smartCommand.Parameters.TryGetValue("appRequired", out var app) ? 
                    app.ToString() : string.Empty;
                
                if (!string.IsNullOrEmpty(appRequired) && appRequired != "system")
                {
                    // Create a launch command for the required application
                    var launchCommand = new GeminiCommand
                    {
                        CommandType = "launch",
                        Target = appRequired,
                        Parameters = new Dictionary<string, object>()
                    };
                    
                    return await _parentService.ExecuteCommandDirectlyAsync(launchCommand);
                }
                
                // If all else fails
                return new CommandResult
                {
                    Success = false,
                    Message = $"I understand you want to: {analysis}, but I'm not sure how to do that yet.",
                    Suggestions = new List<string> 
                    { 
                        "Try being more specific",
                        "Try breaking your request into smaller steps",
                        "Try using a simpler command" 
                    }
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Smart command processing error: {ex.Message}");
                return new CommandResult
                {
                    Success = false,
                    Message = $"Error processing command: {ex.Message}",
                    Suggestions = new List<string> { "Please try a different command" }
                };
            }
        }
        
        /// <summary>
        /// Processes the command with the AI service to understand intent
        /// </summary>
        private async Task<GeminiCommand> ProcessWithAI(string commandText)
        {
            try
            {
                // Format the prompt with the user's command
                string prompt = string.Format(AI_COMMAND_PROMPT, commandText);
                
                // Use the Gemini service to get a structured response
                string response = await _geminiService.SendRawPromptAsync(prompt);
                
                // Extract JSON from the response
                int startIdx = response.IndexOf('{');
                int endIdx = response.LastIndexOf('}');
                
                if (startIdx >= 0 && endIdx >= 0 && endIdx > startIdx)
                {
                    string jsonStr = response.Substring(startIdx, endIdx - startIdx + 1);
                    
                    // Parse the JSON response into a temporary structure
                    var parsedResponse = JsonSerializer.Deserialize<Dictionary<string, object>>(jsonStr);
                    
                    // Create a GeminiCommand with the parameters from the response
                    var result = new GeminiCommand
                    {
                        CommandType = "smart",
                        Target = commandText,
                        Parameters = new Dictionary<string, object>()
                    };
                    
                    // Copy over the parameters from the parsed response
                    foreach (var kvp in parsedResponse)
                    {
                        result.Parameters[kvp.Key] = kvp.Value;
                    }
                    
                    // Try to extract action type
                    if (parsedResponse.TryGetValue("actionType", out var actionType))
                    {
                        result.CommandType = actionType.ToString();
                    }
                    
                    // Try to extract specific action
                    if (parsedResponse.TryGetValue("specificAction", out var specificAction))
                    {
                        result.Action = specificAction.ToString();
                    }
                    
                    return result;
                }
                
                // Default return if JSON parsing fails
                return new GeminiCommand
                {
                    CommandType = "smart",
                    Target = commandText,
                    Parameters = new Dictionary<string, object>
                    {
                        { "analysis", "Failed to parse the response from AI" }
                    }
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"AI processing error: {ex.Message}");
                
                // Return a basic command with the error
                return new GeminiCommand
                {
                    CommandType = "smart",
                    Target = commandText,
                    Parameters = new Dictionary<string, object>
                    {
                        { "analysis", $"Error processing with AI: {ex.Message}" }
                    }
                };
            }
        }
        
        /// <summary>
        /// Converts the AI-processed command to an executable command
        /// </summary>
        private async Task<GeminiCommand> ConvertToExecutableCommand(GeminiCommand smartCommand, string originalCommand)
        {
            try
            {
                // Extract action type
                string actionType = smartCommand.CommandType;
                
                // Default to the original command type if conversion fails
                if (actionType == "smart")
                {
                    if (smartCommand.Parameters.TryGetValue("actionType", out var actionTypeParam))
                    {
                        actionType = actionTypeParam.ToString();
                    }
                }
                
                // Handle different action types
                switch (actionType.ToLower())
                {
                    case "launch":
                        string appName = string.Empty;
                        if (smartCommand.Parameters.TryGetValue("appRequired", out var app))
                        {
                            appName = app.ToString();
                        }
                        
                        return new GeminiCommand
                        {
                            CommandType = "launch",
                            Target = appName,
                            Parameters = ExtractCommandParameters(smartCommand)
                        };
                        
                    case "ui_interact":
                        string targetApp = string.Empty;
                        string element = string.Empty;
                        string action = "click";
                        
                        if (smartCommand.Parameters.TryGetValue("appRequired", out var uiApp))
                        {
                            targetApp = uiApp.ToString();
                        }
                        
                        if (smartCommand.Parameters.TryGetValue("specificAction", out var uiAction))
                        {
                            var actionLower = uiAction.ToString().ToLower();
                            
                            if (actionLower.Contains("click"))
                                action = "click";
                            else if (actionLower.Contains("type") || actionLower.Contains("write") || actionLower.Contains("enter"))
                                action = "type";
                            else if (actionLower.Contains("select"))
                                action = "select";
                            else if (actionLower.Contains("right") && actionLower.Contains("click"))
                                action = "rightclick";
                            else if (actionLower.Contains("double") && actionLower.Contains("click"))
                                action = "doubleclick";
                        }
                        
                        // Extract UI element
                        if (smartCommand.Parameters.TryGetValue("parameters", out var uiParams) && 
                            uiParams is JsonElement jsonElement)
                        {
                            foreach (var prop in jsonElement.EnumerateObject())
                            {
                                if (prop.Name == "element" || prop.Name == "button" || prop.Name == "control" || 
                                    prop.Name == "field" || prop.Name == "target")
                                {
                                    element = prop.Value.ToString();
                                    break;
                                }
                            }
                        }
                        
                        var parameters = new Dictionary<string, object>();
                        
                        // Only add element if it's not empty
                        if (!string.IsNullOrEmpty(element))
                        {
                            parameters["element"] = element;
                        }
                        
                        // Special handling for "write" actions
                        string textContent = ExtractTextParameter(smartCommand);
                        if (action == "type" && string.IsNullOrEmpty(textContent))
                        {
                            // If we couldn't extract text from parameters, try to use the command text itself
                            // This handles cases like "open notepad and write hello irem inside"
                            string lowerOriginalCommand = originalCommand.ToLower();
                            
                            // Check for common text injection patterns in commands
                            string[] writePatterns = {
                                "write ", "type ", "input ", "enter "
                            };
                            
                            foreach (var pattern in writePatterns)
                            {
                                int index = lowerOriginalCommand.IndexOf(pattern);
                                if (index >= 0)
                                {
                                    // Extract everything after the pattern
                                    string extractedText = originalCommand.Substring(index + pattern.Length).Trim();
                                    
                                    // Clean up common suffixes
                                    string[] suffixesToRemove = {
                                        " in it", " inside", " in notepad", " in the notepad", " in document",
                                        " inside it", " to it", " there"
                                    };
                                    
                                    foreach (var suffix in suffixesToRemove)
                                    {
                                        if (extractedText.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                                        {
                                            extractedText = extractedText.Substring(0, extractedText.Length - suffix.Length).Trim();
                                        }
                                    }
                                    
                                    textContent = extractedText;
                                    break;
                                }
                            }
                        }
                        
                        parameters["text"] = textContent;
                        
                        // Create UI command
                        var uiCommand = new GeminiCommand
                        {
                            CommandType = "ui",
                            Action = action,
                            Target = targetApp,
                            Parameters = parameters
                        };
                        
                        try
                        {
                            // Execute UI command
                            var uiResult = await _parentService.ExecuteCommandDirectlyAsync(uiCommand);
                            
                            // If UI automation fails but it was a type action, try direct keyboard input
                            if (!uiResult.Success && action == "type" && !string.IsNullOrEmpty(textContent))
                            {
                                Console.WriteLine("UI automation failed, trying direct keyboard input");
                                // Use direct keyboard input but return the original UI command
                                await TypeTextDirectly(targetApp, textContent);
                                return uiCommand;
                            }
                            
                            // Return the original UI command
                            return uiCommand;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"UI automation error: {ex.Message}");
                            
                            // If it was a type action, try direct keyboard input
                            if (action == "type" && !string.IsNullOrEmpty(textContent))
                            {
                                Console.WriteLine("Using direct keyboard input as fallback");
                                await TypeTextDirectly(targetApp, textContent);
                                return uiCommand;
                            }
                            
                            throw;
                        }
                        
                    case "file_operation":
                        string fileAction = smartCommand.Action;
                        string filePath = string.Empty;
                        string content = string.Empty;
                        
                        if (smartCommand.Parameters.TryGetValue("parameters", out var fileParams) && 
                            fileParams is JsonElement fileJsonElement)
                        {
                            foreach (var prop in fileJsonElement.EnumerateObject())
                            {
                                if (prop.Name == "path" || prop.Name == "filePath" || prop.Name == "file")
                                {
                                    filePath = prop.Value.ToString();
                                }
                                else if (prop.Name == "content" || prop.Name == "text")
                                {
                                    content = prop.Value.ToString();
                                }
                            }
                        }
                        
                        if (fileAction?.ToLower()?.Contains("write") == true || 
                            fileAction?.ToLower()?.Contains("save") == true || 
                            fileAction?.ToLower()?.Contains("create") == true)
                        {
                            return new GeminiCommand
                            {
                                CommandType = "writefile",
                                Target = filePath,
                                Parameters = new Dictionary<string, object>
                                {
                                    { "content", content },
                                    { "append", false }
                                }
                            };
                        }
                        else if (fileAction?.ToLower()?.Contains("read") == true || 
                                 fileAction?.ToLower()?.Contains("open") == true)
                        {
                            return new GeminiCommand
                            {
                                CommandType = "readfile",
                                Target = filePath,
                                Parameters = new Dictionary<string, object>()
                            };
                        }
                        break;
                        
                    case "search":
                        string searchQuery = originalCommand;
                        
                        if (smartCommand.Parameters.TryGetValue("parameters", out var searchParams) && 
                            searchParams is JsonElement searchJsonElement)
                        {
                            foreach (var prop in searchJsonElement.EnumerateObject())
                            {
                                if (prop.Name == "query" || prop.Name == "searchQuery" || prop.Name == "term")
                                {
                                    searchQuery = prop.Value.ToString();
                                    break;
                                }
                            }
                        }
                        
                        return new GeminiCommand
                        {
                            CommandType = "search",
                            Target = searchQuery,
                            Parameters = new Dictionary<string, object>()
                        };
                }
                
                // If we can't convert to a known command type, return null
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Command conversion error: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Extracts command parameters from the smart command
        /// </summary>
        private Dictionary<string, object> ExtractCommandParameters(GeminiCommand smartCommand)
        {
            var result = new Dictionary<string, object>();
            
            // Try to extract nested parameters
            if (smartCommand.Parameters.TryGetValue("parameters", out var parameters) && 
                parameters is JsonElement jsonElement)
            {
                foreach (var prop in jsonElement.EnumerateObject())
                {
                    result[prop.Name] = prop.Value.ToString();
                }
            }
            
            return result;
        }
        
        /// <summary>
        /// Extracts text parameter for UI operations
        /// </summary>
        private string ExtractTextParameter(GeminiCommand smartCommand)
        {
            try
            {
                if (smartCommand.Parameters.TryGetValue("parameters", out var parameters) && 
                    parameters is JsonElement jsonElement)
                {
                    foreach (var prop in jsonElement.EnumerateObject())
                    {
                        if (prop.Name == "text" || prop.Name == "content" || prop.Name == "input")
                        {
                            return prop.Value.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting text parameter: {ex.Message}");
            }
            
            return string.Empty;
        }
        
        /// <summary>
        /// Checks if the command is a question or knowledge query
        /// </summary>
        private bool IsQuestionOrQuery(string command)
        {
            string lowerCommand = command.ToLower().Trim();
            
            // Check if it starts with question words
            if (lowerCommand.StartsWith("what") || 
                lowerCommand.StartsWith("how") || 
                lowerCommand.StartsWith("why") || 
                lowerCommand.StartsWith("when") || 
                lowerCommand.StartsWith("where") || 
                lowerCommand.StartsWith("who") || 
                lowerCommand.StartsWith("which") || 
                lowerCommand.Contains("?"))
            {
                return true;
            }
            
            // Check for query phrases
            string[] queryPhrases = {
                "tell me about", "search for", "look up", "find information", "i want to know",
                "explain", "definition of", "meaning of", "search", "find", "look for"
            };
            
            foreach (var phrase in queryPhrases)
            {
                if (lowerCommand.Contains(phrase))
                {
                    return true;
                }
            }
            
            return false;
        }
        
        /// <summary>
        /// Performs a web search for knowledge queries
        /// </summary>
        private async Task<CommandResult> PerformWebSearch(string query)
        {
            try
            {
                // Create a search command
                var searchCommand = new GeminiCommand
                {
                    CommandType = "search",
                    Target = query,
                    Parameters = new Dictionary<string, object>()
                };
                
                // Use the parent service to execute the search
                return await _parentService.ExecuteCommandDirectlyAsync(searchCommand);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Web search error: {ex.Message}");
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to search for information: {ex.Message}",
                    Suggestions = new List<string> { "Please try a different search term" }
                };
            }
        }

        /// <summary>
        /// Types text using SendKeys as a fallback when UI automation fails
        /// </summary>
        private async Task<CommandResult> TypeTextDirectly(string appName, string text)
        {
            try
            {
                // Give the application time to initialize
                await Task.Delay(1000);
                
                // Directly send keystrokes using SendKeys
                SendKeys.SendWait(text);
                
                return new CommandResult
                {
                    Success = true,
                    Message = $"Typed '{text}' into {appName} using keyboard input"
                };
            }
            catch (Exception ex)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to type text using keyboard input: {ex.Message}"
                };
            }
        }
    }
} 