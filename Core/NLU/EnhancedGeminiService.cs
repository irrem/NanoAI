using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoAI.Core.NLU.Handlers;
using NanoAI.Gemini;

namespace NanoAI.Core.NLU
{
    /// <summary>
    /// Enhanced service for handling various commands including composite commands
    /// </summary>
    public class EnhancedGeminiService : ICommandProcessingService
    {
        private readonly GeminiService _geminiService;
        private readonly List<ICommandHandler> _commandHandlers;

        /// <summary>
        /// Creates an instance of EnhancedGeminiService
        /// </summary>
        /// <param name="geminiService">Base Gemini service</param>
        public EnhancedGeminiService(GeminiService geminiService)
        {
            _geminiService = geminiService ?? throw new ArgumentNullException(nameof(geminiService));
            _commandHandlers = new List<ICommandHandler>();
        }

        /// <summary>
        /// Registers a command handler
        /// </summary>
        /// <param name="handler">Command handler to register</param>
        public void RegisterCommandHandler(ICommandHandler handler)
        {
            if (handler == null)
                throw new ArgumentNullException(nameof(handler));

            _commandHandlers.Add(handler);
            
            // If this is a composite handler, give it a reference to this service
            if (handler is CompositeCommandHandler compositeHandler)
            {
                compositeHandler.ParentService = this;
            }
        }

        /// <summary>
        /// Processes a user command and executes it
        /// </summary>
        /// <param name="userInput">User command input</param>
        /// <returns>Result of the command execution</returns>
        public async Task<CommandResult> ProcessCommandAsync(string userInput)
        {
            try
            {
                // First check if this is a composite command with multiple actions
                var compositeHandler = _commandHandlers.FirstOrDefault(h => h is CompositeCommandHandler) as CompositeCommandHandler;
                if (compositeHandler != null && compositeHandler.CanHandle(new GeminiCommand { Target = userInput }))
                {
                    return await compositeHandler.ExecuteAsync(new GeminiCommand 
                    { 
                        CommandType = "composite",
                        Target = userInput,
                        Parameters = new Dictionary<string, object>()
                    });
                }

                // Process with Gemini API
                var command = await _geminiService.SendPromptAsync(userInput);
                
                if (command == null)
                {
                    return new CommandResult(false, "Failed to process command.");
                }

                // Find handler for this command
                var handler = _commandHandlers.FirstOrDefault(h => h.CanHandle(command));
                
                if (handler != null)
                {
                    return await handler.ExecuteAsync(command);
                }
                else
                {
                    // Try to use the SmartCommandHandler as a fallback
                    var smartHandler = _commandHandlers.FirstOrDefault(h => h is SmartCommandHandler);
                    
                    if (smartHandler != null)
                    {
                        // Convert to a smart command
                        var smartCommand = new GeminiCommand
                        {
                            CommandType = "smart",
                            Target = command.Target ?? userInput,
                            Parameters = command.Parameters ?? new Dictionary<string, object>()
                        };
                        
                        return await smartHandler.ExecuteAsync(smartCommand);
                    }
                    
                    return new CommandResult(false, $"No handler found for command type: {command.CommandType}");
                }
            }
            catch (Exception ex)
            {
                return new CommandResult(false, $"Error processing command: {ex.Message}");
            }
        }

        /// <summary>
        /// Executes a command directly without going through the API
        /// </summary>
        /// <param name="command">The command to execute</param>
        /// <returns>Result of the command execution</returns>
        public async Task<CommandResult> ExecuteCommandDirectlyAsync(GeminiCommand command)
        {
            try
            {
                if (command == null)
                {
                    return new CommandResult(false, "Command is null");
                }

                // Special handling for composite commands
                if (command.CommandType == "composite")
                {
                    var compositeHandler = _commandHandlers.FirstOrDefault(h => h is CompositeCommandHandler) as CompositeCommandHandler;
                    if (compositeHandler != null)
                    {
                        return await compositeHandler.ExecuteAsync(command);
                    }
                }

                // Smart fallback for "custom" commands that are UI-related
                if (command.CommandType.Equals("custom", StringComparison.OrdinalIgnoreCase))
                {
                    // List of UI-related action keywords
                    var uiVerbs = new List<string> {
                        "click", "tap", "press", "push",
                        "select", "choose", "pick",
                        "type", "enter", "input", "write",
                        "drag", "move", "scroll",
                        "right-click", "rightclick", "context menu",
                        "double-click", "doubleclick", "dblclick",
                        "check", "uncheck", "toggle",
                        "screenshot", "capture"
                    };

                    string targetText = command.Target?.ToLower() ?? string.Empty;
                    
                    // Check if any UI verb is present in the target text
                    foreach (var verb in uiVerbs)
                    {
                        if (targetText.StartsWith(verb + " ") || targetText.Contains(" " + verb + " "))
                        {
                            Console.WriteLine($"Rerouting custom command '{targetText}' to UI handler");
                            
                            // Extract parts for the UI command
                            string action = GetUIAction(targetText);
                            string elementName = GetElementName(targetText, action);
                            string appName = GetAppName(targetText) ?? "active";
                            
                            // Create a proper UI command
                            var uiCommand = new GeminiCommand
                            {
                                CommandType = "ui",
                                Action = action,
                                Target = appName,
                                Parameters = new Dictionary<string, object>
                                {
                                    { "element", elementName }
                                }
                            };
                            
                            // Find UI handler
                            var uiHandler = _commandHandlers.FirstOrDefault(h => h.CanHandle(uiCommand));
                            if (uiHandler != null)
                            {
                                return await uiHandler.ExecuteAsync(uiCommand);
                            }
                        }
                    }
                }

                // Find handler for this command
                var handler = _commandHandlers.FirstOrDefault(h => h.CanHandle(command));
                
                if (handler != null)
                {
                    return await handler.ExecuteAsync(command);
                }
                else
                {
                    // Try to use the SmartCommandHandler as a fallback
                    var smartHandler = _commandHandlers.FirstOrDefault(h => h is SmartCommandHandler);
                    
                    if (smartHandler != null)
                    {
                        // Convert to a smart command
                        var smartCommand = new GeminiCommand
                        {
                            CommandType = "smart",
                            Target = command.Target,
                            Parameters = command.Parameters ?? new Dictionary<string, object>()
                        };
                        
                        return await smartHandler.ExecuteAsync(smartCommand);
                    }
                    
                    return new CommandResult(false, $"No handler found for command type: {command.CommandType}");
                }
            }
            catch (Exception ex)
            {
                return new CommandResult(false, $"Error executing command: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts the UI action from a command text
        /// </summary>
        private string GetUIAction(string text)
        {
            // Map common verbs to standard UI actions
            if (text.Contains("click") || text.Contains("tap") || text.Contains("press") || text.Contains("push"))
                return "click";
            if (text.Contains("type") || text.Contains("enter") || text.Contains("input") || text.Contains("write"))
                return "type";
            if (text.Contains("select") || text.Contains("choose") || text.Contains("pick"))
                return "select";
            if (text.Contains("double-click") || text.Contains("doubleclick") || text.Contains("dblclick"))
                return "doubleclick";
            if (text.Contains("right-click") || text.Contains("rightclick") || text.Contains("context menu"))
                return "rightclick";
            if (text.Contains("drag") || text.Contains("move"))
                return "drag";
            if (text.Contains("scroll"))
                return "scroll";
            if (text.Contains("screenshot") || text.Contains("capture"))
                return "screenshot";
            
            // Default to click if no specific action is recognized
            return "click";
        }

        /// <summary>
        /// Extracts the element name from a command text
        /// </summary>
        private string GetElementName(string text, string action)
        {
            try
            {
                // Extract element name based on action
                if (action == "click" || action == "doubleclick" || action == "rightclick")
                {
                    // Format: "click [on] [the] <element> [in <app>]"
                    string pattern = action + " ";
                    int startIdx = text.IndexOf(pattern);
                    if (startIdx >= 0)
                    {
                        string rest = text.Substring(startIdx + pattern.Length).Trim();
                        
                        // Remove optional "on" or "the"
                        if (rest.StartsWith("on "))
                            rest = rest.Substring(3).Trim();
                        if (rest.StartsWith("the "))
                            rest = rest.Substring(4).Trim();
                        
                        // Extract up to "in <app>" if present
                        int inIdx = rest.LastIndexOf(" in ");
                        if (inIdx > 0)
                            rest = rest.Substring(0, inIdx).Trim();
                        
                        // Clean up quotes if present
                        return rest.Trim('\'', '"');
                    }
                }
                else if (action == "type")
                {
                    // Format: "type <text> [in <element> [in <app>]]"
                    if (text.Contains(" in "))
                    {
                        int lastInIdx = text.LastIndexOf(" in ");
                        string possibleElement = text.Substring(lastInIdx + 4).Trim();
                        
                        // Check if there's another "in" for the app name
                        int appInIdx = possibleElement.LastIndexOf(" in ");
                        if (appInIdx > 0)
                            possibleElement = possibleElement.Substring(0, appInIdx).Trim();
                        
                        return possibleElement.Trim('\'', '"');
                    }
                }
            }
            catch { /* Ignore parsing errors and return a generic element name */ }
            
            return "element";
        }

        /// <summary>
        /// Extracts the application name from a command text
        /// </summary>
        private string GetAppName(string text)
        {
            try
            {
                // Look for " in <app>" pattern
                int inIdx = text.LastIndexOf(" in ");
                if (inIdx > 0)
                {
                    string appName = text.Substring(inIdx + 4).Trim();
                    return appName.Trim('\'', '"');
                }
            }
            catch { /* Ignore parsing errors */ }
            
            return null;
        }

        /// <summary>
        /// Gets all registered command handlers
        /// </summary>
        public IEnumerable<ICommandHandler> GetCommandHandlers()
        {
            return _commandHandlers.AsReadOnly();
        }
    }
} 