using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using NanoAI.Gemini;
using System.Linq;

namespace NanoAI.Core.NLU.Handlers
{
    /// <summary>
    /// Handles commands to launch external applications, URLs, or files.
    /// </summary>
    public class LaunchCommandHandler : ICommandHandler
    {
        private readonly Dictionary<string, string> _knownApplications;
        private readonly ICommandProcessingService _parentService;

        public string CommandType => "launch";

        public LaunchCommandHandler(ICommandProcessingService parentService)
        {
            _knownApplications = InitializeKnownApplications();
            _parentService = parentService;
        }

        public bool CanHandle(GeminiCommand command)
        {
            return command.CommandType.Equals("launch", StringComparison.OrdinalIgnoreCase);
        }

        private Dictionary<string, string> InitializeKnownApplications()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // System applications
                { "notepad", "notepad.exe" },
                { "calculator", "calc.exe" },
                { "explorer", "explorer.exe" },
                { "cmd", "cmd.exe" },
                { "powershell", "powershell.exe" },
                { "terminal", "wt.exe" },
                
                // Browsers
                { "chrome", @"C:\Program Files\Google\Chrome\Application\chrome.exe" },
                { "firefox", @"C:\Program Files\Mozilla Firefox\firefox.exe" },
                { "edge", @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" },
                
                // Office applications
                { "word", @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" },
                { "excel", @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" },
                { "powerpoint", @"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" },
                { "outlook", @"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" },
                
                // Development tools
                { "vscode", @"C:\Users\$USERNAME\AppData\Local\Programs\Microsoft VS Code\Code.exe" },
                { "visualstudio", @"C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe" }
            };
        }

        public async Task<CommandResult> ExecuteAsync(GeminiCommand command)
        {
            if (string.IsNullOrEmpty(command.Target))
            {
                return new CommandResult
                {
                    Success = false,
                    Message = "No application or command sequence provided",
                    Suggestions = new List<string> { "Please provide an application to launch" }
                };
            }
            
            // Check if this is a command sequence with "then" keyword
            if (command.Target.Contains(" then "))
            {
                var commands = SplitIntoCommandSequence(command.Target);
                
                if (commands == null || commands.Count == 0)
                {
                    return new CommandResult
                    {
                        Success = false,
                        Message = "Could not parse command sequence",
                        Suggestions = new List<string> { "Please check your command syntax" }
                    };
                }
                
                var results = new List<string>();
                var allErrors = new List<string>();
                var anySuccess = false;
                
                // Execute each command in sequence
                for (int i = 0; i < commands.Count; i++)
                {
                    var result = await _parentService.ProcessCommandAsync(commands[i]);
                    
                    if (result.Success)
                    {
                        results.Add($"Step {i + 1}: {result.Message}");
                        anySuccess = true;
                    }
                    else
                    {
                        results.Add($"Step {i + 1} failed: {result.Message}");
                        allErrors.Add(result.Message);
                    }
                }
                
                return new CommandResult
                {
                    Success = anySuccess,
                    Message = string.Join("\n", results),
                    Suggestions = allErrors
                };
            }
            else
            {
                // This is a direct launch command, not a sequence
                string appName = command.Target.Trim();
                
                // Check if it's a URL
                if (IsUrl(appName))
                {
                    return await LaunchUrl(appName);
                }
                
                // Check if it's a local file path
                if (File.Exists(appName))
                {
                    return await LaunchFile(appName);
                }
                
                // Otherwise, treat as an application
                return await LaunchApplication(appName);
            }
        }

        private List<string> SplitIntoCommandSequence(string input)
        {
            var commands = new List<string>();
            var parts = input.Split(new[] { " then ", " and " }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var part in parts)
            {
                if (!string.IsNullOrWhiteSpace(part))
                {
                    commands.Add(part.Trim());
                }
            }
            
            return commands;
        }

        private GeminiCommand ParseCommandLocally(string input)
        {
            string lowerInput = input.ToLower().Trim();
            
            // Launch/start/open/run commands
            if (lowerInput.StartsWith("launch ") || lowerInput.StartsWith("start ") || lowerInput.StartsWith("open ") || lowerInput.StartsWith("run "))
            {
                int spaceIndex = input.IndexOf(' ');
                string appName = spaceIndex >= 0 ? input.Substring(spaceIndex + 1).Trim() : string.Empty;
                
                return new GeminiCommand
                {
                    CommandType = "launch",
                    Target = appName,
                    Parameters = new Dictionary<string, object>()
                };
            }
            
            // Write commands
            if (lowerInput.StartsWith("write "))
            {
                string content = input.Substring(6).Trim();
                
                return new GeminiCommand
                {
                    CommandType = "ui",
                    Action = "type",
                    Target = "notepad", // Default to Notepad
                    Parameters = new Dictionary<string, object>
                    {
                        { "text", content }
                    }
                };
            }
            
            return null;
        }

        private bool IsUrl(string target)
        {
            return Uri.TryCreate(target, UriKind.Absolute, out Uri uriResult)
                && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
        }

        private async Task<CommandResult> LaunchUrl(string url)
        {
            // Ensure URL has a protocol prefix
            if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                url = "https://" + url;
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                };

                Process.Start(psi);
                Console.WriteLine($"Launched URL: {url}");
                return new CommandResult { Success = true, Message = $"Launched URL: {url}" };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to launch URL '{url}': {ex.Message}");
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to launch URL '{url}': {ex.Message}"
                };
            }
        }

        private async Task<CommandResult> LaunchFile(string filePath)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };

                Process.Start(psi);
                Console.WriteLine($"Launched file: {filePath}");
                return new CommandResult { Success = true, Message = $"Launched file: {filePath}" };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to launch file '{filePath}': {ex.Message}");
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to launch file '{filePath}': {ex.Message}"
                };
            }
        }

        private async Task<CommandResult> LaunchApplication(string appName)
        {
            string appPath = null;

            // Print debug information showing what we're trying to launch
            Console.WriteLine($"Attempting to launch application: '{appName}' (preserving original case)");

            // Check if it's a known application
            if (_knownApplications.TryGetValue(appName, out string path))
            {
                appPath = path;
                
                // Replace $USERNAME with actual username if needed
                if (appPath.Contains("$USERNAME"))
                {
                    string username = Environment.UserName;
                    appPath = appPath.Replace("$USERNAME", username);
                }
                Console.WriteLine($"Found in known applications: {appPath}");
            }
            else
            {
                // Assume the appName is the executable name; let Shell search PATHEXT
                appPath = appName;
                // If it has no extension, add .exe for explicit search
                if (string.IsNullOrEmpty(Path.GetExtension(appPath)))
                {
                    appPath = appName + ".exe";
                }
                Console.WriteLine($"Not in known apps, trying with path: {appPath}");
                
                // Enhanced search in multiple locations
                string foundPath = SearchForApplication(appPath);
                if (!string.IsNullOrEmpty(foundPath))
                {
                    appPath = foundPath;
                    Console.WriteLine($"Found application at: {appPath}");
                }
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = appPath,
                    UseShellExecute = true
                };

                Process.Start(psi);
                Console.WriteLine($"Launched application: {appPath}");
                return new CommandResult { Success = true, Message = $"Launched application: {appName}" };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to launch application '{appName}' (path: {appPath}): {ex.Message}");
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to launch application '{appName}': {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Enhanced search for an application in various system and common directories
        /// </summary>
        private string SearchForApplication(string appName)
        {
            // List of common directories to search for applications
            var searchPaths = new List<string>
            {
                // Application's base directory
                AppContext.BaseDirectory,
                
                // Current directory
                Directory.GetCurrentDirectory(),
                
                // System directories
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                Environment.GetFolderPath(Environment.SpecialFolder.SystemX86),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                
                // Common application locations
                @"C:\Program Files",
                @"C:\Program Files (x86)",

                // Desktop and user directories
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            };
            
            // First check Start Menu shortcuts - this is often how applications are launched
            string startMenuPath = SearchStartMenuShortcuts(appName);
            if (!string.IsNullOrEmpty(startMenuPath))
                return startMenuPath;
            
            // Search for exact match in each directory
            foreach (var dir in searchPaths)
            {
                if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
                    continue;
                
                string fullPath = Path.Combine(dir, appName);
                Console.WriteLine($"Checking: {fullPath}");
                if (File.Exists(fullPath))
                    return fullPath;
                
                // Also try searching one level of subdirectories
                try
                {
                    foreach (var subDir in Directory.GetDirectories(dir))
                    {
                        string subPath = Path.Combine(subDir, appName);
                        if (File.Exists(subPath))
                            return subPath;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error searching directory {dir}: {ex.Message}");
                }
            }
            
            // If nothing found, just return the original app name to let the shell try to find it
            return null;
        }

        /// <summary>
        /// Searches Start Menu for shortcuts (.lnk files) that match the application name
        /// </summary>
        private string SearchStartMenuShortcuts(string appName)
        {
            try
            {
                // Get both common and user-specific Start Menu directories
                string[] startMenuFolders = new[]
                {
                    Environment.GetFolderPath(Environment.SpecialFolder.StartMenu),
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs")
                };
                
                Console.WriteLine("Searching Start Menu for application shortcuts...");
                
                string appNameWithoutExt = Path.GetFileNameWithoutExtension(appName);
                
                foreach (var folder in startMenuFolders)
                {
                    if (!Directory.Exists(folder))
                        continue;
                        
                    // Search for shortcuts with similar names (handles partial matches)
                    var shortcuts = Directory.GetFiles(folder, "*.lnk", SearchOption.AllDirectories)
                        .Where(f => Path.GetFileNameWithoutExtension(f).IndexOf(appNameWithoutExt, StringComparison.OrdinalIgnoreCase) >= 0)
                        .ToList();
                    
                    if (shortcuts.Count > 0)
                    {
                        Console.WriteLine($"Found {shortcuts.Count} possible matches in Start Menu:");
                        
                        foreach (var shortcut in shortcuts)
                        {
                            Console.WriteLine($"  Shortcut: {shortcut}");
                            // For shortcut resolution, we'll just return the .lnk file itself
                            // Windows shell execution can handle .lnk files directly
                            return shortcut;
                        }
                    }
                }
                
                Console.WriteLine("No matching shortcuts found in Start Menu");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error searching Start Menu: {ex.Message}");
            }
            
            return null;
        }

        /// <summary>
        /// Extracts the text to write from a command
        /// </summary>
        private string GetTextFromCommand(string command)
        {
            string lowerCommand = command.ToLower().Trim();
            string text = "";
            
            if (lowerCommand.StartsWith("write "))
                text = command.Substring(6).Trim();
            else if (lowerCommand.StartsWith("type "))
                text = command.Substring(5).Trim();
            else if (lowerCommand.StartsWith("input "))
                text = command.Substring(6).Trim();
            else if (lowerCommand.StartsWith("enter "))
                text = command.Substring(6).Trim();
                
            // Clean up common suffixes 
            string[] suffixesToRemove = {
                " in it", " inside", " in notepad", " in the notepad", " in document",
                " inside it", " to it", " there"
            };
                
            foreach (var suffix in suffixesToRemove)
            {
                if (text.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    text = text.Substring(0, text.Length - suffix.Length).Trim();
                }
            }
            
            return text;
        }
    }

    public interface ICommandProcessingService
    {
        Task<CommandResult> ProcessCommandAsync(string input);
        Task<CommandResult> ExecuteCommandDirectlyAsync(GeminiCommand command);
    }
} 