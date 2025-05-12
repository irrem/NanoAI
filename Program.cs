using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using NanoAI.Core.NLU;
using NanoAI.Core.NLU.Handlers;
using NanoAI.Gemini;
using System.Collections.Generic;

namespace NanoAI
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Set the culture to English
            CultureInfo.CurrentCulture = new CultureInfo("en-US");
            CultureInfo.CurrentUICulture = new CultureInfo("en-US");
            
            Console.WriteLine("Starting NanoAI...");
            
            // Get API key from environment variable
            string apiKey = Environment.GetEnvironmentVariable("GEMINI_API_KEY");
            apiKey = "AIzaSyAH0_maE6Jx8-ciKHACZjUxYUcoME_JA8Q";
            // If not found in environment, try to read from a file
            if (string.IsNullOrEmpty(apiKey))
            {
                string apiKeyFile = "apikey.txt";
                if (File.Exists(apiKeyFile))
                {
                    try
                    {
                        ///   apiKey = File.ReadAllText(apiKeyFile).Trim();
                         apiKey = "AIzaSyAH0_maE6Jx8-ciKHACZjUxYUcoME_JA8Q";
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error reading API key file: {ex.Message}");
                    }
                }
            }
            
            if (string.IsNullOrEmpty(apiKey))
            {
                Console.WriteLine("Error: Gemini API key not found. Please set the GEMINI_API_KEY environment variable or create an apikey.txt file");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return;
            }

            // Initialize Gemini service
            var geminiService = new GeminiService(apiKey);
            
            // Initialize enhanced service with NLU capabilities
            var enhancedService = new EnhancedGeminiService(geminiService);
            
            // Register command handlers
            enhancedService.RegisterCommandHandler(new LaunchCommandHandler(enhancedService));
            enhancedService.RegisterCommandHandler(new ProjectCommandHandler());
            enhancedService.RegisterCommandHandler(new UIActionHandler());
            enhancedService.RegisterCommandHandler(new CompositeCommandHandler(enhancedService));
            enhancedService.RegisterCommandHandler(new ServiceControlHandler());
            enhancedService.RegisterCommandHandler(new WriteFileCommandHandler());
            enhancedService.RegisterCommandHandler(new SearchCommandHandler(geminiService));
            enhancedService.RegisterCommandHandler(new SmartCommandHandler(geminiService, enhancedService));

            // Show example commands
            Console.WriteLine("\nExample commands:");
            Console.WriteLine("- start notepad");
            Console.WriteLine("- open calculator");
            Console.WriteLine("- run my python script");
            Console.WriteLine("- launch NDIntegrationServer then start notepad");
            Console.WriteLine("- Click 'Save' button in Notepad");
            Console.WriteLine("- Take screenshot of Calculator");
            Console.WriteLine("- service target=wuauserv action=status");
            Console.WriteLine("- open notepad and write hello world");
            Console.WriteLine("- what is the capital of France?");
            Console.WriteLine("- help me create a PowerPoint presentation");
            Console.WriteLine("- open Excel and create a budget spreadsheet");
            Console.WriteLine("\nType 'exit' to quit");
            
            // Main interaction loop
            while (true)
            {
                Console.Write("\n> ");
                string? input = Console.ReadLine();
                
                if (string.IsNullOrEmpty(input) || input.ToLower() == "exit")
                {
                    break;
                }
                
                try
                {
                    // Local parsing to bypass Gemini API issues
                    var command = ParseCommandLocally(input);
                    var result = command != null ? 
                        await enhancedService.ExecuteCommandDirectlyAsync(command) : 
                        await enhancedService.ProcessCommandAsync(input);
                        
                    Console.WriteLine(result.Message);
                    
                    if (result.Suggestions?.Count > 0)
                    {
                        Console.WriteLine("\nSuggestions:");
                        foreach (var suggestion in result.Suggestions)
                        {
                            Console.WriteLine($"- {suggestion}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }
            
            Console.WriteLine("Goodbye!");
        }
        
        // Simple local command parser for basic commands
        private static GeminiCommand ParseCommandLocally(string input)
        {
            string lowerInput = input.ToLower().Trim();
            
            // For composite commands
            if (lowerInput.Contains(" then ") || lowerInput.Contains(" and "))
            {
                // Split the command but preserve case for each part
                string separator = lowerInput.Contains(" then ") ? " then " : " and ";
                int separatorIndex = input.IndexOf(separator, StringComparison.OrdinalIgnoreCase);
                var parts = new string[] { 
                    input.Substring(0, separatorIndex), 
                    input.Substring(separatorIndex + separator.Length) 
                };
                
                if (parts.Length >= 2)
                {
                    return new GeminiCommand
                    {
                        CommandType = "composite",
                        Target = input,
                        Parameters = new Dictionary<string, object>()
                    };
                }
            }
            
            // Launch/start/open/run commands
            if (lowerInput.StartsWith("launch ") || lowerInput.StartsWith("start ") || lowerInput.StartsWith("open ") || lowerInput.StartsWith("run "))
            {
                // Use original input to preserve casing
                int spaceIndex = input.IndexOf(' ');
                string appName = spaceIndex >= 0 ? input.Substring(spaceIndex + 1).Trim() : string.Empty;
                
                return new GeminiCommand
                {
                    CommandType = "launch",
                    Target = appName,
                    Parameters = new Dictionary<string, object>()
                };
            }
            
            // Write/Save commands
            if (lowerInput.StartsWith("write ") || lowerInput.StartsWith("save "))
            {
                string content = "";
                string filePath = "";
                bool foundSaveAs = false;
                
                // Get original case input
                string originalInput = input.Trim();
                
                if (lowerInput.Contains(" as "))
                {
                    // Format: "write|save [text] as [filename]"
                    int asIndex = lowerInput.IndexOf(" as ");
                    content = originalInput.Substring(originalInput.IndexOf(' ') + 1, asIndex - originalInput.IndexOf(' ') - 1).Trim();
                    filePath = originalInput.Substring(asIndex + 4).Trim();
                    foundSaveAs = true;
                }
                else if (lowerInput.Contains(" to "))
                {
                    // Format: "write|save [text] to [filename]"
                    int toIndex = lowerInput.IndexOf(" to ");
                    content = originalInput.Substring(originalInput.IndexOf(' ') + 1, toIndex - originalInput.IndexOf(' ') - 1).Trim();
                    filePath = originalInput.Substring(toIndex + 4).Trim();
                    foundSaveAs = true;
                }
                else
                {
                    // Just "write [text]" - use default filename
                    content = originalInput.Substring(originalInput.IndexOf(' ') + 1).Trim();
                    filePath = "output.txt";
                }
                
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
            
            // UI commands
            if (lowerInput.StartsWith("click "))
            {
                string rest = lowerInput.Substring(6).Trim();
                string element = rest;
                string app = "notepad"; // Default
                
                if (rest.Contains(" in "))
                {
                    var parts = rest.Split(new[] { " in " }, StringSplitOptions.None);
                    element = parts[0].Trim();
                    app = parts[1].Trim();
                }
                
                return new GeminiCommand
                {
                    CommandType = "ui",
                    Action = "click",
                    Target = app,
                    Parameters = new Dictionary<string, object>
                    {
                        { "element", element }
                    }
                };
            }
            
            if (lowerInput.StartsWith("take screenshot"))
            {
                string app = "desktop";
                if (lowerInput.Contains(" of "))
                {
                    app = lowerInput.Substring(lowerInput.IndexOf(" of ") + 4).Trim();
                }
                
                return new GeminiCommand
                {
                    CommandType = "ui",
                    Action = "screenshot",
                    Target = app,
                    Parameters = new Dictionary<string, object>()
                };
            }
            
            // Service commands
            if (lowerInput.StartsWith("service "))
            {
                string rest = lowerInput.Substring(8).Trim();
                string target = "";
                string action = "status";
                
                if (rest.Contains("target="))
                {
                    int targetStart = rest.IndexOf("target=") + 7;
                    int targetEnd = rest.IndexOf(" ", targetStart);
                    if (targetEnd == -1) targetEnd = rest.Length;
                    target = rest.Substring(targetStart, targetEnd - targetStart);
                }
                
                if (rest.Contains("action="))
                {
                    int actionStart = rest.IndexOf("action=") + 7;
                    int actionEnd = rest.IndexOf(" ", actionStart);
                    if (actionEnd == -1) actionEnd = rest.Length;
                    action = rest.Substring(actionStart, actionEnd - actionStart);
                }
                
                return new GeminiCommand
                {
                    CommandType = "service",
                    Action = action,
                    Target = target,
                    Parameters = new Dictionary<string, object>()
                };
            }
            
            return null;
        }
    }
    
    /// <summary>
    /// Handles commands to write content to files
    /// </summary>
    public class WriteFileCommandHandler : ICommandHandler
    {
        public string CommandType => "writefile";

        public bool CanHandle(GeminiCommand command)
        {
            return command.CommandType.Equals("writefile", StringComparison.OrdinalIgnoreCase);
        }

        public async Task<CommandResult> ExecuteAsync(GeminiCommand command)
        {
            if (command == null)
            {
                throw new ArgumentNullException(nameof(command));
            }

            string target = command.Target;
            string content = command.GetParameterValue("content") ?? string.Empty;
            bool append = command.GetParameterValue("append")?.ToLower() == "true";

            if (string.IsNullOrEmpty(target))
            {
                return new CommandResult
                {
                    Success = false,
                    Message = "No file path specified for writing",
                    Suggestions = new List<string> { "Please specify a file path to write to." }
                };
            }

            try
            {
                // Determine if this is a relative or absolute path
                string fullPath = target;
                if (!Path.IsPathRooted(target))
                {
                    // Check for special folder references
                    if (target.StartsWith("desktop", StringComparison.OrdinalIgnoreCase))
                    {
                        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        fullPath = Path.Combine(desktopPath, target.Substring(7).TrimStart('/', '\\'));
                    }
                    else if (target.StartsWith("documents", StringComparison.OrdinalIgnoreCase))
                    {
                        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        fullPath = Path.Combine(documentsPath, target.Substring(9).TrimStart('/', '\\'));
                    }
                    else
                    {
                        // Use current directory as default
                        fullPath = Path.Combine(Directory.GetCurrentDirectory(), target);
                    }
                }

                // Ensure directory exists
                string directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Write to file
                if (append)
                {
                    File.AppendAllText(fullPath, content);
                    Console.WriteLine($"Appended text to file: {fullPath}");
                }
                else
                {
                    File.WriteAllText(fullPath, content);
                    Console.WriteLine($"Wrote text to file: {fullPath}");
                }

                return new CommandResult
                {
                    Success = true,
                    Message = append ? 
                        $"Successfully appended to file: {fullPath}" : 
                        $"Successfully wrote to file: {fullPath}"
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to file '{target}': {ex.Message}");
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to write to file: {ex.Message}",
                    Suggestions = new List<string> {
                        "Check that the directory exists",
                        "Check that you have permission to write to this location"
                    }
                };
            }
        }
    }
}
