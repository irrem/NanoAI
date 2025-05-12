using System;
using System.Collections.Generic;
using System.Net.Http;
// using System.Net.Http.Json; - Bu referansı kaldırıyorum
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace NanoAI.Gemini
{
    /// <summary>
    /// Service that communicates with Google Gemini AI API
    /// </summary>
    public class GeminiService
    {
        private readonly HttpClient _httpClient;
        public readonly string _apiKey;
        public readonly string _apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent";

        /// <summary>
        /// Creates a GeminiService instance
        /// </summary>
        /// <param name="apiKey">Google API key</param>
        public GeminiService(string apiKey)
        {
            _apiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Sends a request to Gemini API and gets the response
        /// </summary>
        /// <param name="prompt">User request</param>
        /// <returns>Command object</returns>
        public async Task<GeminiCommand> SendPromptAsync(string prompt)
        {
            if (string.IsNullOrEmpty(prompt))
            {
                throw new ArgumentException("Request cannot be empty", nameof(prompt));
            }

            try
            {
                // Create prompt with system instructions
                string fullPrompt = $@"You are a computer command assistant. You should convert natural language commands from users into machine-processable commands.

You support the following command types:
- launch: Launches an application (Target: application name)
- close: Closes an application (Target: application name)
- type: Types text (Target: application to type in or empty, Parameters: {{""text"": ""text to type""}})
- click: Clicks a UI element (Target: element name to click, Parameters: {{""application"": ""application name"", ""x"": x_coordinate, ""y"": y_coordinate}})
- search: Searches the web (Target: search query, Parameters: {{""engine"": ""search engine""}})
- research: Researches a topic (Target: research topic, Parameters: {{""outputFile"": ""result file""}})
- readfile: Reads a file (Target: file path)
- writefile: Writes to a file (Target: file path, Parameters: {{""content"": ""content"", ""append"": true/false}})
- runscript: Runs a script (Target: script file, Parameters: {{""arguments"": ""arguments""}})
- systeminfo: Gets system information (Target: info type - ""os"", ""cpu"", ""memory"", ""disk"", ""user"", ""running"")
- servicecontrol: Controls Windows services (Target: service name, Parameters: {{""action"": ""start/stop/restart/status""}})
- project: Runs a project or script (Target: project name, Parameters: {{""runAsAdmin"": true/false}})
- ui: Performs UI automation actions (Target: action type - ""click"", ""type"", ""select"", ""rightclick"", ""doubleclick"", ""drag"", ""scroll"", ""gettext"", ""wait"", ""screenshot"", Parameters: specific to action type)
- custom: Other commands (Target: command text)

For UI automation, use these parameters:
- click: Parameters: {{""application"": ""app name"", ""element"": ""element name"", ""type"": ""button/textbox/etc"", ""x"": x_coord, ""y"": y_coord}}
- type: Parameters: {{""application"": ""app name"", ""element"": ""element name"", ""text"": ""text to type""}}
- select: Parameters: {{""application"": ""app name"", ""element"": ""element name"", ""item"": ""item to select""}}
- rightclick, doubleclick: Same parameters as click
- drag: Parameters: {{""application"": ""app name"", ""source"": ""element name"", ""target"": ""element name"", ""sourceX"": x1, ""sourceY"": y1, ""targetX"": x2, ""targetY"": y2}}
- scroll: Parameters: {{""application"": ""app name"", ""element"": ""element name"", ""direction"": ""up/down/left/right"", ""amount"": lines}}
- gettext: Parameters: {{""application"": ""app name"", ""element"": ""element name""}}
- wait: Parameters: {{""application"": ""app name"", ""element"": ""element name"", ""timeout"": seconds}}
- screenshot: Parameters: {{""application"": ""app name"", ""path"": ""file path"", ""fullscreen"": true/false}}

Analyze the user command and determine the appropriate command type, target, and parameters. Respond in the following JSON format:

{{
  ""commandType"": ""[command type]"",
  ""target"": ""[target]"",
  ""parameters"": {{
    ""parameter1"": ""value1"",
    ""parameter2"": ""value2""
  }}
}}

User input: {prompt}";

                // Create Gemini API request
                var request = new
                {
                    contents = new[]
                    {
                        new
                        {
                            role = "user",
                            parts = new[]
                            {
                                new
                                {
                                    text = fullPrompt
                                }
                            }
                        }
                    }
                };

                string jsonContent = JsonSerializer.Serialize(request);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync($"{_apiUrl}?key={_apiKey}", content);
                
                if (!response.IsSuccessStatusCode)
                {
                    if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                    {
                        // If API endpoint not found, try to parse the command directly
                        return new GeminiCommand
                        {
                            CommandType = "custom",
                            Target = prompt,
                            Parameters = new Dictionary<string, object>()
                        };
                    }
                    response.EnsureSuccessStatusCode();
                }

                var responseContent = await response.Content.ReadAsStringAsync();
                var parsedResponse = JsonDocument.Parse(responseContent);

                // Extract text from Gemini API response structure
                var responseText = parsedResponse
                    .RootElement
                    .GetProperty("candidates")[0]
                    .GetProperty("content")
                    .GetProperty("parts")[0]
                    .GetProperty("text")
                    .GetString();

                // Find JSON part in response
                int startIdx = responseText.IndexOf('{');
                int endIdx = responseText.LastIndexOf('}');
                
                if (startIdx >= 0 && endIdx >= 0 && endIdx > startIdx)
                {
                    string jsonStr = responseText.Substring(startIdx, endIdx - startIdx + 1);
                    
                    // Convert JSON to command object
                    var command = JsonSerializer.Deserialize<GeminiCommand>(jsonStr);
                    
                    // Create empty dictionary if Parameters is null
                    if (command.Parameters == null)
                    {
                        command.Parameters = new Dictionary<string, object>();
                    }
                    
                    return command;
                }
                
                // If JSON not found
                throw new Exception("Could not get a valid command. API response: " + responseText);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Command processing error: {ex.Message}");
                
                // Try to parse basic commands locally
                return ParseBasicCommandsLocally(prompt);
            }
        }
        
        /// <summary>
        /// Parses basic commands locally when the API fails
        /// </summary>
        private GeminiCommand ParseBasicCommandsLocally(string prompt)
        {
            string lowerPrompt = prompt.ToLower().Trim();
            
            // Handle basic application launch commands
            if (lowerPrompt.StartsWith("open ") || lowerPrompt.StartsWith("launch ") || 
                lowerPrompt.StartsWith("start ") || lowerPrompt.StartsWith("run "))
            {
                string appName = "";
                if (lowerPrompt.StartsWith("open ")) appName = lowerPrompt.Substring(5).Trim();
                else if (lowerPrompt.StartsWith("launch ")) appName = lowerPrompt.Substring(7).Trim();
                else if (lowerPrompt.StartsWith("start ")) appName = lowerPrompt.Substring(6).Trim();
                else if (lowerPrompt.StartsWith("run ")) appName = lowerPrompt.Substring(4).Trim();
                
                // Extract first word as application name
                if (!string.IsNullOrEmpty(appName))
                {
                    string[] parts = appName.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                    {
                        return new GeminiCommand
                        {
                            CommandType = "launch",
                            Target = parts[0],
                            Parameters = new Dictionary<string, object>()
                        };
                    }
                }
            }
            
            // Handle basic close commands
            if (lowerPrompt.StartsWith("close ") || lowerPrompt.StartsWith("quit ") || 
                lowerPrompt.StartsWith("exit ") || lowerPrompt.StartsWith("stop "))
            {
                string appName = "";
                if (lowerPrompt.StartsWith("close ")) appName = lowerPrompt.Substring(6).Trim();
                else if (lowerPrompt.StartsWith("quit ")) appName = lowerPrompt.Substring(5).Trim();
                else if (lowerPrompt.StartsWith("exit ")) appName = lowerPrompt.Substring(5).Trim();
                else if (lowerPrompt.StartsWith("stop ")) appName = lowerPrompt.Substring(5).Trim();
                
                if (!string.IsNullOrEmpty(appName))
                {
                    string[] parts = appName.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                    {
                        return new GeminiCommand
                        {
                            CommandType = "close",
                            Target = parts[0],
                            Parameters = new Dictionary<string, object>()
                        };
                    }
                }
            }
            
            // Handle click commands
            if (lowerPrompt.StartsWith("click "))
            {
                string elementName = lowerPrompt.Substring(6).Trim();
                if (!string.IsNullOrEmpty(elementName))
                {
                    // Check if there's an "in" or "on" to specify the application
                    string appName = null;
                    int inIndex = elementName.LastIndexOf(" in ");
                    int onIndex = elementName.LastIndexOf(" on ");
                    
                    if (inIndex > 0)
                    {
                        appName = elementName.Substring(inIndex + 4).Trim();
                        elementName = elementName.Substring(0, inIndex).Trim();
                    }
                    else if (onIndex > 0)
                    {
                        appName = elementName.Substring(onIndex + 4).Trim();
                        elementName = elementName.Substring(0, onIndex).Trim();
                    }
                    
                    // Remove quotes around element name if present
                    if (elementName.StartsWith("'") && elementName.EndsWith("'"))
                    {
                        elementName = elementName.Substring(1, elementName.Length - 2);
                    }
                    else if (elementName.StartsWith("\"") && elementName.EndsWith("\""))
                    {
                        elementName = elementName.Substring(1, elementName.Length - 2);
                    }
                    
                    var parameters = new Dictionary<string, object>();
                    if (!string.IsNullOrEmpty(appName))
                    {
                        parameters.Add("application", appName);
                    }
                    parameters.Add("element", elementName);
                    
                    return new GeminiCommand
                    {
                        CommandType = "ui",
                        Target = "click",
                        Parameters = parameters
                    };
                }
            }
            
            // Handle type commands
            if (lowerPrompt.StartsWith("type "))
            {
                string text = lowerPrompt.Substring(5).Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    // Check if there's an "in" to specify the application or element
                    string appName = null;
                    string elementName = null;
                    int inIndex = text.LastIndexOf(" in ");
                    
                    if (inIndex > 0)
                    {
                        string target = text.Substring(inIndex + 4).Trim();
                        text = text.Substring(0, inIndex).Trim();
                        
                        // If target contains quotes, it's likely an element name
                        if (target.Contains("'") || target.Contains("\""))
                        {
                            elementName = target.Trim('\'', '"');
                        }
                        else
                        {
                            appName = target;
                        }
                    }
                    
                    var parameters = new Dictionary<string, object>();
                    if (!string.IsNullOrEmpty(appName))
                    {
                        parameters.Add("application", appName);
                    }
                    if (!string.IsNullOrEmpty(elementName))
                    {
                        parameters.Add("element", elementName);
                    }
                    parameters.Add("text", text);
                    
                    return new GeminiCommand
                    {
                        CommandType = "ui",
                        Target = "type",
                        Parameters = parameters
                    };
                }
            }
            
            // Handle screenshot commands
            if (lowerPrompt.StartsWith("screenshot") || lowerPrompt.StartsWith("take screenshot"))
            {
                string appName = null;
                
                // Extract "of [appname]" if present
                if (lowerPrompt.Contains(" of "))
                {
                    int ofIndex = lowerPrompt.IndexOf(" of ");
                    appName = lowerPrompt.Substring(ofIndex + 4).Trim();
                }
                
                var parameters = new Dictionary<string, object>();
                if (!string.IsNullOrEmpty(appName))
                {
                    parameters.Add("application", appName);
                }
                
                return new GeminiCommand
                {
                    CommandType = "ui",
                    Target = "screenshot",
                    Parameters = parameters
                };
            }
            
            // Default to custom command
            return new GeminiCommand
            {
                CommandType = "custom",
                Target = prompt,
                Parameters = new Dictionary<string, object>()
            };
        }
        
        /// <summary>
        /// Sends a raw prompt to Gemini API and gets the text response
        /// </summary>
        public async Task<string> SendRawPromptAsync(string prompt)
        {
            if (string.IsNullOrEmpty(prompt))
            {
                throw new ArgumentException("Prompt cannot be empty", nameof(prompt));
            }

            try
            {
                var request = new
                {
                    contents = new[]
                    {
                        new
                        {
                            role = "user",
                            parts = new[]
                            {
                                new { text = prompt }
                            }
                        }
                    }
                };

                string jsonContent = JsonSerializer.Serialize(request);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync($"{_apiUrl}?key={_apiKey}", content);
                response.EnsureSuccessStatusCode();

                var responseContent = await response.Content.ReadAsStringAsync();
                var parsedResponse = JsonDocument.Parse(responseContent);

                return parsedResponse
                    .RootElement
                    .GetProperty("candidates")[0]
                    .GetProperty("content")
                    .GetProperty("parts")[0]
                    .GetProperty("text")
                    .GetString();
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
    }
} 