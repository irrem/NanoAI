using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using NanoAI.Gemini;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

namespace NanoAI
{
    /// <summary>
    /// Class that processes user commands and performs appropriate actions
    /// </summary>
    public class CommandProcessor
    {
        private readonly GeminiService _geminiService;
        private readonly Dictionary<string, Application> _runningApplications;
        private readonly HttpClient _httpClient;

        /// <summary>
        /// Creates a CommandProcessor instance
        /// </summary>
        /// <param name="apiKey">Gemini API key</param>
        public CommandProcessor(string apiKey)
        {
            _geminiService = new GeminiService(apiKey);
            _runningApplications = new Dictionary<string, Application>();
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Processes user command and performs appropriate action
        /// </summary>
        /// <param name="userInput">User input</param>
        /// <returns>Processing result</returns>
        public async Task<string> ProcessCommandAsync(string userInput)
        {
            try
            {
                // Sequence delimiters for multi-step commands
                string[] sequenceDelimiters = new[] { 
                    " then ", " and then ", " after that ", " later ", ", then ", 
                    " and ", " , ", ", ", "; ", ";" 
                };
                
                // Check command sequence
                string[] commandSequence = SplitIntoCommandSequence(userInput, sequenceDelimiters);
                
                // If there are multiple commands, process them sequentially
                if (commandSequence.Length > 1)
                {
                    return await ProcessSequentialCommandsAsync(commandSequence);
                }
                
                // Single command processing continues
                // Check for special instructions in the user input first
                bool runAsAdmin = false;
                
                // Check if it's a service control request
                if (userInput.ToLower().Contains("service"))
                {
                    // Process service command
                    return await ControlServiceAsync(userInput);
                }
                
                // Look for "as administrator" phrase
                string[] adminKeywords = new[] { 
                    "as administrator", "as admin", "elevated", "with administrative rights", 
                    "with admin privileges", "with administrator permissions",
                    "run as admin", "launch as admin", "start as admin",
                    "execute as admin", "open as admin"
                };
                
                foreach (var keyword in adminKeywords)
                {
                    if (userInput.ToLower().Contains(keyword))
                    {
                        runAsAdmin = true;
                        // Remove keyword to preserve main part of request
                        userInput = userInput.Replace(keyword, "").Trim();
                        break;
                    }
                }
                
                // Analyze command through Gemini API
                var command = await _geminiService.SendPromptAsync(userInput);
                
                Console.WriteLine($"Command detected: {command.CommandType} - {command.Target}");
                
                // Add run as admin parameter
                if (runAsAdmin)
                {
                    command.Parameters["runAsAdmin"] = true;
                }
                
                // Process based on command type
                switch (command.CommandType.ToLower())
                {
                    case "launch":
                    case "open":
                    case "start":
                    case "run":
                    case "execute":
                        return LaunchApplication(command);
                    case "close":
                    case "exit":
                    case "quit":
                    case "stop":
                    case "end":
                        return CloseApplication(command);
                    case "type":
                    case "input":
                    case "enter":
                        return TypeText(command);
                    case "click":
                    case "press":
                    case "select":
                        return ClickElement(command);
                    case "search":
                    case "find":
                        return await SearchAsync(command);
                    case "research":
                    case "investigate":
                        return await ResearchAsync(command);
                    case "readfile":
                    case "read":
                        return ReadFile(command);
                    case "writefile":
                    case "write":
                        return WriteFile(command);
                    case "runscript":
                    case "script":
                        return RunScript(command);
                    case "systeminfo":
                    case "system":
                        return GetSystemInfo(command);
                    case "servicecontrol":
                    case "service":
                        return ControlService(command);
                    case "custom":
                        return await ProcessCustomCommandAsync(command);
                    default:
                        return $"Unsupported command type: {command.CommandType}";
                }
            }
            catch (Exception ex)
            {
                return $"Command processing error: {ex.Message}";
            }
        }
        
        /// <summary>
        /// Splits user command into sequential steps
        /// </summary>
        private string[] SplitIntoCommandSequence(string input, string[] delimiters)
        {
            // Make the sentence easier to process
            string normalizedInput = input.Trim();
            
            // Find the most suitable delimiter and split sequentially
            List<string> commandParts = new List<string>();
            string currentInput = normalizedInput;
            
            foreach (var delimiter in delimiters)
            {
                if (currentInput.Contains(delimiter))
                {
                    string[] parts = currentInput.Split(new[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);
                    
                    // Add first part to list, rejoin the rest
                    if (parts.Length > 0)
                    {
                        commandParts.Add(parts[0].Trim());
                        
                        // Rejoin remaining parts (for other delimiters)
                        if (parts.Length > 1)
                        {
                            currentInput = string.Join(" ", parts.Skip(1)).Trim();
                        }
                        else
                        {
                            currentInput = "";
                        }
                    }
                    
                    // If nothing remains, exit loop
                    if (string.IsNullOrWhiteSpace(currentInput))
                        break;
                }
            }
            
            // If there's still unprocessed text, add it as the last command
            if (!string.IsNullOrWhiteSpace(currentInput))
            {
                commandParts.Add(currentInput.Trim());
            }
            
            // If no commands found, return original input as a single command
            if (commandParts.Count == 0)
            {
                commandParts.Add(normalizedInput);
            }
            
            return commandParts.ToArray();
        }
        
        /// <summary>
        /// Processes a sequence of commands
        /// </summary>
        private async Task<string> ProcessSequentialCommandsAsync(string[] commandSequence)
        {
            StringBuilder result = new StringBuilder("Executing sequential commands:\n");
            string lastAppLaunched = ""; // Track the last launched application
            
            for (int i = 0; i < commandSequence.Length; i++)
            {
                string command = commandSequence[i];
                result.AppendLine($"Step {i+1}/{commandSequence.Length}: {command}");
                
                try
                {
                    // Detect "Check if running" or similar check
                    if (IsCheckRunningCommand(command))
                    {
                        if (!string.IsNullOrEmpty(lastAppLaunched))
                        {
                            string appName = ExtractApplicationName(lastAppLaunched);
                            bool isRunning = IsApplicationRunning(appName);
                            
                            result.AppendLine($"Application status check: {appName} {(isRunning ? "is running ✓" : "is not running ✗")}");
                            
                            if (isRunning)
                            {
                                result.AppendLine($"➤ Application {appName} is running successfully.");
                            }
                            else
                            {
                                result.AppendLine($"⚠️ Warning: Application {appName} is not running or cannot be found!");
                                
                                // Suggest alternative solution
                                result.AppendLine($"Suggestions: Try launching the application manually or specify the full file path.");
                            }
                        }
                        else
                        {
                            result.AppendLine("No previously launched application found to check.");
                        }
                        
                        if (i < commandSequence.Length - 1)
                        {
                            result.AppendLine("Moving to the next step...");
                        }
                        continue;
                    }
                    
                    // Normal command processing
                    string stepResult = await ProcessCommandAsync(command);
                    result.AppendLine($"Result: {stepResult}");
                    
                    // Error check - if a step fails, suggest alternatives
                    if (stepResult.Contains("error") || stepResult.Contains("not found") || stepResult.Contains("failed"))
                    {
                        // If it's an application launch error and the application wasn't found, suggest similar apps
                        if (stepResult.Contains("not found") && command.ToLower().Contains("launch"))
                        {
                            string appName = ExtractApplicationName(command);
                            List<string> similarApps = FindSimilarApplications(appName);
                            
                            if (similarApps.Count > 0)
                            {
                                result.AppendLine("The following similar applications were found:");
                                foreach (var app in similarApps.Take(5))
                                {
                                    result.AppendLine($"➤ {app}");
                                }
                                result.AppendLine("Would you like to launch one of these applications?");
                            }
                        }
                        
                        result.AppendLine($"Step {i+1} failed, stopping sequential processing.");
                        break;
                    }
                    
                    // After application launch operations, short wait period
                    if (stepResult.Contains("launched"))
                    {
                        // Track the launched application
                        lastAppLaunched = command;
                        
                        // Wait 2 seconds for application to start
                        await Task.Delay(2000);
                        
                        // Check if application is running
                        if (command.ToLower().Contains("launch") || command.ToLower().Contains("start") || command.ToLower().Contains("open"))
                        {
                            // Try to extract application name
                            string appName = ExtractApplicationName(command);
                            if (!string.IsNullOrEmpty(appName))
                            {
                                bool isRunning = IsApplicationRunning(appName);
                                result.AppendLine($"Application status check: {appName} {(isRunning ? "is running ✓" : "is not running ✗")}");
                                
                                if (!isRunning)
                                {
                                    result.AppendLine($"⚠️ Warning: {appName} was launched but does not appear to be running!");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.AppendLine($"Error: {ex.Message}");
                    result.AppendLine("Sequential processing stopped due to error.");
                    break;
                }
                
                // If not the last step, add transition message
                if (i < commandSequence.Length - 1)
                {
                    result.AppendLine("Moving to the next step...");
                }
            }
            
            result.AppendLine("Sequential command processing completed.");
            return result.ToString();
        }
        
        /// <summary>
        /// Finds similar applications to the specified application name
        /// </summary>
        private List<string> FindSimilarApplications(string appName)
        {
            List<string> similarApps = new List<string>();
            try
            {
                if (string.IsNullOrEmpty(appName))
                    return similarApps;
                
                // Pre-processing: catch dots, dashes, underscores and spaces
                string normalizedAppName = appName.ToLower().Replace(".", " ").Replace("-", " ").Replace("_", " ");
                string[] searchTerms = normalizedAppName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                
                // Similarity threshold
                int threshold = 5; // Levenshtein distance
                
                // First search in Windows start menu (shortcuts are usually here)
                string[] startMenuPaths = new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs")
                };
                
                foreach (string startMenuPath in startMenuPaths)
                {
                    if (!Directory.Exists(startMenuPath))
                        continue;
                    
                    try
                    {
                        var shortcutFiles = Directory.GetFiles(startMenuPath, "*.lnk", SearchOption.AllDirectories);
                        foreach (string shortcutFile in shortcutFiles)
                        {
                            string shortcutName = Path.GetFileNameWithoutExtension(shortcutFile).ToLower();
                            
                            // Search by terms
                            bool containsAllTerms = true;
                            foreach (var term in searchTerms)
                            {
                                if (term.Length <= 2) continue; // Skip very short terms (a, an, etc.)
                                if (!shortcutName.Contains(term))
                                {
                                    containsAllTerms = false;
                                    break;
                                }
                            }
                            
                            if (containsAllTerms || searchTerms.Any(term => shortcutName.Contains(term)) || 
                                LevenshteinDistance(shortcutName, normalizedAppName.Replace(" ", "")) <= threshold)
                            {
                                string targetPath = ResolveShortcutTarget(shortcutFile);
                                if (!string.IsNullOrEmpty(targetPath) && File.Exists(targetPath))
                                {
                                    string appDisplayName = Path.GetFileNameWithoutExtension(targetPath);
                                    similarApps.Add(appDisplayName);
                                }
                                else
                                {
                                    // Shortcut could not be resolved but name is similar, add shortcut name
                                    similarApps.Add(Path.GetFileNameWithoutExtension(shortcutFile));
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Ignore access errors
                    }
                }
                
                // Search in Program Files folders
                string[] programFolders = new[]
                {
                    @"C:\Program Files", 
                    @"C:\Program Files (x86)",
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
                };
                
                // Find folders containing each word
                foreach (var folder in programFolders)
                {
                    if (!Directory.Exists(folder))
                        continue;
                        
                    try
                    {
                        // Find .exe files at 2 levels depth
                        foreach (var dir in Directory.GetDirectories(folder))
                        {
                            string dirName = Path.GetFileName(dir).ToLower();
                            
                            // Search for directories containing the search terms
                            bool directoryMatches = searchTerms.Any(term => dirName.Contains(term));
                            
                            if (directoryMatches || searchTerms.Length == 1) // If only one term, match all directories
                            {
                                try
                                {
                                    // Check .exe files at 1st level
                                    foreach (var exe in Directory.GetFiles(dir, "*.exe"))
                                    {
                                        string exeName = Path.GetFileNameWithoutExtension(exe).ToLower();
                                        
                                        // Search by terms
                                        bool matchesTerms = searchTerms.Any(term => 
                                            exeName.Contains(term) || 
                                            LevenshteinDistance(exeName, term) <= threshold);
                                            
                                        if (matchesTerms || LevenshteinDistance(exeName, normalizedAppName.Replace(" ", "")) <= threshold)
                                        {
                                            similarApps.Add(Path.GetFileNameWithoutExtension(exe));
                                        }
                                    }
                                    
                                    // Check .exe files at 2nd level
                                    foreach (var subdir in Directory.GetDirectories(dir))
                                    {
                                        try
                                        {
                                            string subdirName = Path.GetFileName(subdir).ToLower();
                                            bool subdirMatches = searchTerms.Any(term => subdirName.Contains(term));
                                            
                                            if (directoryMatches || subdirMatches)
                                            {
                                                foreach (var exe in Directory.GetFiles(subdir, "*.exe"))
                                                {
                                                    string exeName = Path.GetFileNameWithoutExtension(exe).ToLower();
                                                    
                                                    // Search by terms
                                                    bool matchesTerms = searchTerms.Any(term => 
                                                        exeName.Contains(term) || 
                                                        LevenshteinDistance(exeName, term) <= threshold);
                                                        
                                                    if (matchesTerms || LevenshteinDistance(exeName, normalizedAppName.Replace(" ", "")) <= threshold)
                                                    {
                                                        similarApps.Add(Path.GetFileNameWithoutExtension(exe));
                                                    }
                                                }
                                            }
                                        }
                                        catch { /* Ignore access errors */ }
                                    }
                                }
                                catch { /* Ignore access errors */ }
                            }
                        }
                    }
                    catch { /* Ignore access errors */ }
                }
                
                // Search in current directory
                try
                {
                    foreach (var exe in Directory.GetFiles(Environment.CurrentDirectory, "*.exe", SearchOption.AllDirectories))
                    {
                        string exeName = Path.GetFileNameWithoutExtension(exe).ToLower();
                        
                        // Search by terms
                        bool matchesTerms = searchTerms.Any(term => 
                            exeName.Contains(term) || 
                            LevenshteinDistance(exeName, term) <= threshold);
                            
                        if (matchesTerms || LevenshteinDistance(exeName, normalizedAppName.Replace(" ", "")) <= threshold)
                        {
                            similarApps.Add(Path.GetFileNameWithoutExtension(exe));
                        }
                    }
                }
                catch { /* Ignore errors */ }
                
                // Search user folders
                string[] userFolders = new[]
                {
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                };
                
                foreach (var folder in userFolders)
                {
                    if (!Directory.Exists(folder))
                        continue;
                        
                    try
                    {
                        foreach (var exe in Directory.GetFiles(folder, "*.exe", SearchOption.AllDirectories))
                        {
                            string exeName = Path.GetFileNameWithoutExtension(exe).ToLower();
                            
                            // Search by terms
                            bool matchesTerms = searchTerms.Any(term => 
                                exeName.Contains(term) || 
                                LevenshteinDistance(exeName, term) <= threshold);
                                
                            if (matchesTerms || LevenshteinDistance(exeName, normalizedAppName.Replace(" ", "")) <= threshold)
                            {
                                similarApps.Add(Path.GetFileNameWithoutExtension(exe));
                            }
                        }
                    }
                    catch { /* Ignore access errors */ }
                }
                
                // Special cases: NDIS and Web Server
                if (normalizedAppName.Contains("ndis") && normalizedAppName.Contains("web") && normalizedAppName.Contains("server"))
                {
                    // Check common application names related to NDIS and Web Server
                    string[] possibilities = new[]
                    {
                        "NDISServer", 
                        "NDIS.Web.Server", 
                        "NDISweb", 
                        "WebServer",
                        "NDIntegrationServer", 
                        "NDServer", 
                        "NDHostingService",
                        "NDIS.Core.Server"
                    };
                    
                    foreach (var possibility in possibilities)
                    {
                        var foundExes = SearchForExactApplication(possibility);
                        foreach (var exe in foundExes)
                        {
                            similarApps.Add(Path.GetFileNameWithoutExtension(exe));
                        }
                    }
                }
                
                // List installed applications (from Process.GetProcesses() names)
                try
                {
                    var allProcesses = Process.GetProcesses();
                    var matchingProcesses = allProcesses
                        .Where(p => {
                            string procName = p.ProcessName.ToLower();
                            return searchTerms.Any(term => 
                                    procName.Contains(term)) || 
                                    LevenshteinDistance(procName, normalizedAppName.Replace(" ", "")) <= threshold;
                        })
                        .Select(p => p.ProcessName)
                        .ToList();
                        
                    similarApps.AddRange(matchingProcesses);
                }
                catch { /* Ignore errors */ }
                
                // Remove duplicates, sort, and filter out empty entries
                return similarApps
                    .Where(app => !string.IsNullOrWhiteSpace(app))
                    .Distinct()
                    .OrderBy(a => a)
                    .ToList();
            }
            catch
            {
                return similarApps;
            }
        }
        
        /// <summary>
        /// Searches for an exact application match
        /// </summary>
        private List<string> SearchForExactApplication(string exactAppName)
        {
            List<string> foundPaths = new List<string>();
            
            // .exe extension added
            string appNameWithExe = exactAppName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) 
                ? exactAppName 
                : exactAppName + ".exe";
                
            try
            {
                // Where command used to search in PATH
                var whereProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "where",
                        Arguments = appNameWithExe,
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                
                whereProcess.Start();
                string whereOutput = whereProcess.StandardOutput.ReadToEnd();
                whereProcess.WaitForExit();
                
                if (!string.IsNullOrEmpty(whereOutput))
                {
                    var paths = whereOutput.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                    foundPaths.AddRange(paths);
                }
            }
            catch { /* Ignore errors */ }
            
            // Quick search in Program Files folders
            string[] programFolders = new[]
            {
                @"C:\Program Files", 
                @"C:\Program Files (x86)",
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                Environment.CurrentDirectory // Current directory
            };
            
            foreach (var folder in programFolders)
            {
                if (!Directory.Exists(folder))
                    continue;
                
                try
                {
                    var files = Directory.GetFiles(folder, appNameWithExe, SearchOption.AllDirectories);
                    foundPaths.AddRange(files);
                }
                catch { /* Ignore errors */ }
            }
            
            return foundPaths.Distinct().ToList();
        }
        
        /// <summary>
        /// Checks if the command is a check running command
        /// </summary>
        private bool IsCheckRunningCommand(string command)
        {
            string normalizedCommand = command.ToLower().Trim();
            
            string[] checkKeywords = new[]
            {
                "çalışıyor mu", "çalıştığını kontrol et", "durumunu kontrol et", 
                "kontrol et", "çalışıp çalışmadığını kontrol et", 
                "başladı mı", "açıldı mı", "hazır mı"
            };
            
            foreach (var keyword in checkKeywords)
            {
                if (normalizedCommand.Contains(keyword))
                {
                    return true;
                }
            }
            
            return false;
        }
        
        /// <summary>
        /// Extracts application name from command
        /// </summary>
        private string ExtractApplicationName(string command)
        {
            try
            {
                // Find the text after the start commands
                string[] startCommands = { "başlat", "çalıştır", "aç", "start", "run", "launch", "open" };
                
                foreach (var startCmd in startCommands)
                {
                    if (command.ToLower().Contains(startCmd))
                    {
                        int index = command.ToLower().IndexOf(startCmd) + startCmd.Length;
                        string remaining = command.Substring(index).Trim();
                        
                        // Take the first word (it could be the application name)
                        string appName = remaining.Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)[0];
                        return appName;
                    }
                }
                
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Checks if the application is running
        /// </summary>
        private bool IsApplicationRunning(string appName)
        {
            try
            {
                // Remove .exe extension
                if (appName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    appName = Path.GetFileNameWithoutExtension(appName);
                }
                
                // Check running processes
                var processes = Process.GetProcessesByName(appName);
                
                // Check for similar names as well
                if (processes.Length == 0)
                {
                    // Get all processes and filter out similar ones
                    var allProcesses = Process.GetProcesses();
                    processes = allProcesses.Where(p => 
                        p.ProcessName.ToLower().Contains(appName.ToLower()) || 
                        (p.MainWindowTitle != null && p.MainWindowTitle.ToLower().Contains(appName.ToLower())))
                        .ToArray();
                }
                
                return processes.Length > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Launches the application
        /// </summary>
        private string LaunchApplication(GeminiCommand command)
        {
            string appName = command.Target;
            
            if (string.IsNullOrEmpty(appName))
            {
                return "No application name specified to launch.";
            }

            try
            {
                // Check "as administrator" phrase
                bool runAsAdmin = false;
                if (command.Parameters.TryGetValue("runAsAdmin", out var adminValue))
                {
                    if (adminValue is bool boolValue)
                        runAsAdmin = boolValue;
                    else if (adminValue?.ToString()?.ToLower() == "true")
                        runAsAdmin = true;
                }

                if (_runningApplications.ContainsKey(appName))
                {
                    return $"{appName} is already running.";
                }

                // Start application based on its type
                Process process = null;
                var startInfo = new ProcessStartInfo();
                
                // Get arguments
                string arguments = command.Parameters.TryGetValue("arguments", out var args) ? args?.ToString() : string.Empty;
                
                switch (appName.ToLower())
                {
                    case "calculator":
                    case "hesap makinesi":
                        startInfo.FileName = "calc.exe";
                        break;
                    case "notepad":
                    case "not defteri":
                        startInfo.FileName = "notepad.exe";
                        break;
                    default:
                        // If the file name is not a full path, try to find it in the system first
                        if (!File.Exists(appName) && !appName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            // Search the system for the application
                            string foundPath = FindApplication(appName);
                            if (!string.IsNullOrEmpty(foundPath))
                            {
                                startInfo.FileName = foundPath;
                            }
                            else
                            {
                                // If the application couldn't be found, suggest similar ones
                                List<string> similarApps = FindSimilarApplications(appName);
                                if (similarApps.Count > 0)
                                {
                                    StringBuilder response = new StringBuilder();
                                    response.AppendLine($"Application '{appName}' not found. Similar applications:");
                                    
                                    foreach (var app in similarApps.Take(5)) // Show top 5 matches
                                    {
                                        response.AppendLine($"➤ {app}");
                                    }
                                    
                                    response.AppendLine("To run one of these applications, use: 'open <application name>' command.");
                                    return response.ToString();
                                }
                                
                                // Directly use the exe name
                                startInfo.FileName = appName;
                            }
                        }
                        else
                        {
                            // Directly use the exe name
                            startInfo.FileName = appName;
                        }
                        break;
                }
                
                // Set admin execution settings
                if (runAsAdmin)
                {
                    startInfo.UseShellExecute = true;
                    startInfo.Verb = "runas"; // This means "Run as administrator"
                    Console.WriteLine("Application will be run as administrator");
                }
                else
                {
                    startInfo.UseShellExecute = true;
                }
                
                // Add arguments
                startInfo.Arguments = arguments;
                
                try
                {
                    process = Process.Start(startInfo);
                }
                catch (System.ComponentModel.Win32Exception ex)
                {
                    // File not found error
                    if (ex.NativeErrorCode == 2 || ex.Message.Contains("not found") || ex.Message.Contains("find"))
                    {
                        // Find similar applications
                        List<string> similarApps = FindSimilarApplications(appName);
                        if (similarApps.Count > 0)
                        {
                            StringBuilder response = new StringBuilder();
                            response.AppendLine($"You can try these similar applications instead of '{appName}':");
                            
                            foreach (var app in similarApps.Take(5)) // Show top 5 matches
                            {
                                response.AppendLine($"➤ {app}");
                            }
                        }
                    }
                    
                    // User canceled "Run as administrator" dialog
                    if (ex.NativeErrorCode == 1223) // ERROR_CANCELLED
                    {
                        return "Administrator execution request was canceled.";
                    }
                    throw; // Throw other errors
                }

                if (process != null)
                {
                    try
                    {
                        // FlaUI might not be able to connect when running as admin, handle this case
                        var appInstance = Application.Attach(process);
                        _runningApplications[appName] = appInstance;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Application started but could not connect for automation: {ex.Message}");
                    }
                    return $"{appName} launched" + 
                           (runAsAdmin ? " (as administrator)" : "") + 
                           (string.IsNullOrEmpty(arguments) ? "." : $" (arguments: {arguments})");
                }
                
                return $"Could not launch {appName}.";
            }
            catch (Exception ex)
            {
                // General error handling
                StringBuilder response = new StringBuilder();
                response.AppendLine($"Application launch error: {ex.Message}");
                
                // If the application was not found, suggest similar ones
                if (ex.Message.Contains("not found") || ex.Message.Contains("find"))
                {
                    // Find similar applications
                    List<string> similarApps = FindSimilarApplications(appName);
                    if (similarApps.Count > 0)
                    {
                        response.AppendLine($"You can try these similar applications instead of '{appName}':");
                        
                        foreach (var app in similarApps.Take(5)) // Show top 5 matches
                        {
                            response.AppendLine($"➤ {app}");
                        }
                    }
                }
                
                return response.ToString();
            }
        }

        /// <summary>
        /// Searches for the application name in the system and returns the full path
        /// </summary>
        /// <param name="appName">Application name to search for</param>
        /// <returns>Full path of the application or an empty string if not found</returns>
        private string FindApplication(string appName)
        {
            try
            {
                // Karakter düzeltmesi
                appName = appName.Trim();
                
                // .exe uzantısını ekleyerek de ara
                string appNameWithExe = appName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) 
                    ? appName 
                    : appName + ".exe";
                
                Console.WriteLine($"Sistem üzerinde {appNameWithExe} aranıyor...");

                // 1. Doğrudan yolu dene (eğer verilmişse)
                if (File.Exists(appName))
                {
                    Console.WriteLine($"Uygulama doğrudan bulundu: {appName}");
                    return appName;
                }
                if (File.Exists(appNameWithExe))
                {
                    Console.WriteLine($"Uygulama doğrudan bulundu: {appNameWithExe}");
                    return appNameWithExe;
                }
                
                // 2. Önce where komutunu kullanarak ara (PATH içinde arar)
                var whereProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "where",
                        Arguments = appNameWithExe,
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                
                whereProcess.Start();
                string whereOutput = whereProcess.StandardOutput.ReadToEnd();
                whereProcess.WaitForExit();
                
                if (!string.IsNullOrEmpty(whereOutput))
                {
                    string foundPath = whereOutput.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    if (!string.IsNullOrEmpty(foundPath) && File.Exists(foundPath))
                    {
                        Console.WriteLine($"Uygulama where komutu ile bulundu: {foundPath}");
                        return foundPath;
                    }
                }
                
                // 3. Kullanıcının masaüstü ve belgeleri
                string[] personalFolders = new[]
                {
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                    Environment.CurrentDirectory // Şu anki çalışma dizini
                };
                
                foreach (string folder in personalFolders)
                {
                    if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
                        continue;
                        
                    try
                    {
                        // Tam eşleşme ara
                        string exactPath = Path.Combine(folder, appNameWithExe);
                        if (File.Exists(exactPath))
                        {
                            Console.WriteLine($"Uygulama kişisel klasörde bulundu: {exactPath}");
                            return exactPath;
                        }
                        
                        // Birinci seviye klasörleri kontrol et (çok derin aramaz)
                        foreach (string subFolder in Directory.GetDirectories(folder))
                        {
                            string subPath = Path.Combine(subFolder, appNameWithExe);
                            if (File.Exists(subPath))
                            {
                                Console.WriteLine($"Uygulama alt klasörde bulundu: {subPath}");
                                return subPath;
                            }
                        }
                    }
                    catch
                    {
                        // Erişim reddedilirse sessizce devam et
                    }
                }
                
                // 4. Windows Başlat Menüsü klasörlerini ara
                string[] startMenuPaths = new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs")
                };
                
                foreach (string startMenuPath in startMenuPaths)
                {
                    if (!Directory.Exists(startMenuPath))
                        continue;
                        
                    Console.WriteLine($"Başlat menüsü klasörü aranıyor: {startMenuPath}");
                    
                    try
                    {
                        // .lnk (shortcut) dosyalarını ara
                        var shortcutFiles = Directory.GetFiles(startMenuPath, "*.lnk", SearchOption.AllDirectories);
                        foreach (string shortcutFile in shortcutFiles)
                        {
                            string filename = Path.GetFileNameWithoutExtension(shortcutFile);
                            if (filename.ToLower().Contains(appName.ToLower()))
                            {
                                // Kısayol bulundu, hedef dosyayı bul
                                string targetPath = ResolveShortcutTarget(shortcutFile);
                                if (!string.IsNullOrEmpty(targetPath) && File.Exists(targetPath))
                                {
                                    Console.WriteLine($"Uygulama Başlat menüsü kısayolunda bulundu: {targetPath}");
                                    return targetPath;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Erişim reddedilirse sessizce devam et
                    }
                }
                
                // 5. Yaygın uygulama klasörlerini kontrol et
                string[] commonAppPaths = new[]
                {
                    @"C:\Program Files",
                    @"C:\Program Files (x86)",
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
                };
                
                List<string> foundFiles = new List<string>();
                
                // Önce yaygın uygulama klasörlerinde ara
                foreach (string basePath in commonAppPaths)
                {
                    if (string.IsNullOrEmpty(basePath) || !Directory.Exists(basePath))
                        continue;
                    
                    try
                    {
                        // Birinci ve ikinci seviye klasörlerde ara (derine inmeden)
                        foreach (string dir in Directory.GetDirectories(basePath))
                        {
                            try
                            {
                                // Birinci seviye
                                string exactPath = Path.Combine(dir, appNameWithExe);
                                if (File.Exists(exactPath))
                                {
                                    Console.WriteLine($"Uygulama program klasöründe bulundu: {exactPath}");
                                    return exactPath;
                                }
                                
                                // İsim benzerliğiyle klasörde ara
                                if (Path.GetFileName(dir).ToLower().Contains(appName.ToLower()))
                                {
                                    // Benzer isimli klasörde .exe dosyalarını ara
                                    foreach (string exeFile in Directory.GetFiles(dir, "*.exe"))
                                    {
                                        foundFiles.Add(exeFile);
                                    }
                                }
                                
                                // İkinci seviye
                                foreach (string subdir in Directory.GetDirectories(dir))
                                {
                                    try
                                    {
                                        string subExactPath = Path.Combine(subdir, appNameWithExe);
                                        if (File.Exists(subExactPath))
                                        {
                                            Console.WriteLine($"Uygulama alt program klasöründe bulundu: {subExactPath}");
                                            return subExactPath;
                                        }
                                        
                                        // İkinci seviyedeki benzer isimli klasörlerde ara
                                        if (Path.GetFileName(subdir).ToLower().Contains(appName.ToLower()))
                                        {
                                            foreach (string exeFile in Directory.GetFiles(subdir, "*.exe"))
                                            {
                                                foundFiles.Add(exeFile);
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        // Erişim reddedilirse sessizce devam et
                                    }
                                }
                            }
                            catch
                            {
                                // Erişim reddedilirse sessizce devam et
                            }
                        }
                    }
                    catch
                    {
                        // Erişim reddedilirse sessizce devam et
                    }
                }
                
                // 6. Diğer klasörleri kontrol et (daha az yaygın yerler)
                string[] otherCommonPaths = new[]
                {
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86),
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                };
                
                foreach (string basePath in otherCommonPaths)
                {
                    if (string.IsNullOrEmpty(basePath) || !Directory.Exists(basePath))
                        continue;
                    
                    try
                    {
                        // İsim benzerliğiyle klasör ara
                        var matchingDirs = Directory.GetDirectories(basePath, "*" + appName + "*", SearchOption.TopDirectoryOnly);
                        foreach (string dir in matchingDirs)
                        {
                            try
                            {
                                // Benzer isimli klasörde exe'leri ara
                                var exeFiles = Directory.GetFiles(dir, "*.exe");
                                foundFiles.AddRange(exeFiles);
                            }
                            catch
                            {
                                // Erişim reddedilirse sessizce devam et
                            }
                        }
                    }
                    catch
                    {
                        // Erişim reddedilirse sessizce devam et
                    }
                }
                
                // 7. Bulunan dosyaları değerlendir
                if (foundFiles.Count > 0)
                {
                    // Sonuçları benzerliklere göre sırala
                    foundFiles = foundFiles.Distinct().ToList();
                    
                    // İsim benzerliğine göre en iyi eşleşmeyi bul
                    var bestMatches = foundFiles
                        .Select(f => new 
                        { 
                            Path = f, 
                            Filename = Path.GetFileNameWithoutExtension(f),
                            Score = LevenshteinDistance(Path.GetFileNameWithoutExtension(f).ToLower(), appName.ToLower())
                        })
                        .OrderBy(item => item.Score)
                        .ThenBy(item => item.Filename.Length) // Daha kısa isimleri tercih et
                        .ToList();
                    
                    if (bestMatches.Count > 0)
                    {
                        string bestMatch = bestMatches[0].Path;
                        Console.WriteLine($"Benzerliğe göre en iyi eşleşme bulundu: {bestMatch}");
                        return bestMatch;
                    }
                }
                
                Console.WriteLine($"Uygulama bulunamadı: {appNameWithExe}");
                return string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Uygulama arama hatası: {ex.Message}");
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Resolves the target path of a Windows .lnk shortcut
        /// </summary>
        private string ResolveShortcutTarget(string shortcutPath)
        {
            try
            {
                // Doğrudan okuma - ikili dosyadan okuma
                using (FileStream stream = new FileStream(shortcutPath, FileMode.Open, FileAccess.Read))
                {
                    using (BinaryReader reader = new BinaryReader(stream))
                    {
                        stream.Seek(0x14, SeekOrigin.Begin);
                        int flags = reader.ReadInt32();
                        
                        if ((flags & 1) == 1)
                        {
                            // Klasör kısayolu
                            stream.Seek(0x4C, SeekOrigin.Begin);
                        }
                        else
                        {
                            // Dosya kısayolu
                            stream.Seek(0x3A, SeekOrigin.Begin);
                        }
                        
                        int length = reader.ReadInt16();
                        stream.Seek(0x56, SeekOrigin.Begin);
                        
                        StringBuilder targetPath = new StringBuilder();
                        for (int i = 0; i < length; i++)
                        {
                            char c = reader.ReadChar();
                            if (c != 0) targetPath.Append(c);
                        }
                        
                        return targetPath.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Kısayol çözümleme hatası: {ex.Message}");
                
                // Alternatif yöntem - PowerShell ile çözümleme dene
                try
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-Command \"(New-Object -ComObject WScript.Shell).CreateShortcut('{shortcutPath}').TargetPath\"",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    
                    var process = Process.Start(startInfo);
                    string output = process.StandardOutput.ReadToEnd().Trim();
                    process.WaitForExit();
                    
                    if (!string.IsNullOrEmpty(output))
                    {
                        return output;
                    }
                }
                catch
                {
                    // Powershell de başarısız olduysa boş döndür
                }
                
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Calculates the Levenshtein distance between two strings
        /// Smaller values indicate more similar strings
        /// </summary>
        private int LevenshteinDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];
            
            if (n == 0) return m;
            if (m == 0) return n;
            
            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;
            
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            
            return d[n, m];
        }

        /// <summary>
        /// Closes the application
        /// </summary>
        private string CloseApplication(GeminiCommand command)
        {
            string appName = command.Target;
            
            if (string.IsNullOrEmpty(appName))
            {
                return "Kapatılacak uygulama adı belirtilmedi.";
            }

            try
            {
                if (_runningApplications.TryGetValue(appName, out var app))
                {
                    app.Close();
                    _runningApplications.Remove(appName);
                    return $"{appName} kapatıldı.";
                }
                else
                {
                    // Doğrudan işlem adıyla da kapatmayı dene
                    var processes = Process.GetProcessesByName(appName);
                    if (processes.Length > 0)
                    {
                        foreach (var proc in processes)
                        {
                            proc.Kill();
                        }
                        return $"{processes.Length} adet {appName} işlemi kapatıldı.";
                    }
                }
                
                return $"{appName} çalışan uygulamalar arasında bulunamadı.";
            }
            catch (Exception ex)
            {
                return $"Uygulama kapatma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Types text
        /// </summary>
        private string TypeText(GeminiCommand command)
        {
            string text = command.Parameters.TryGetValue("text", out var value) ? value?.ToString() : null;
            
            if (string.IsNullOrEmpty(text))
            {
                return "Yazılacak metin belirtilmedi.";
            }

            try
            {
                // Eğer hedef belirtilmişse o uygulamaya odaklan
                if (!string.IsNullOrEmpty(command.Target) && 
                    _runningApplications.TryGetValue(command.Target, out var targetApp))
                {
                    using (var automation = new UIA3Automation())
                    {
                        var window = targetApp.GetMainWindow(automation);
                        window.Focus();
                    }
                }
                else if (!string.IsNullOrEmpty(command.Target))
                {
                    // Uygulama adıyla da deneme yap
                    var processes = Process.GetProcessesByName(command.Target);
                    if (processes.Length > 0)
                    {
                        var processApp = Application.Attach(processes[0]);
                        using (var automation = new UIA3Automation())
                        {
                            var window = processApp.GetMainWindow(automation);
                            window.Focus();
                        }
                    }
                }

                // Metni yaz
                Keyboard.Type(text);
                return $"'{text}' metni yazıldı.";
            }
            catch (Exception ex)
            {
                return $"Metin yazma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Clicks on a UI element
        /// </summary>
        private string ClickElement(GeminiCommand command)
        {
            if (string.IsNullOrEmpty(command.Target))
            {
                return "Tıklanacak hedef belirtilmedi.";
            }

            // Koordinatlar belirtilmişse doğrudan koordinata tıkla
            if (command.Parameters.TryGetValue("x", out var xStr) && 
                command.Parameters.TryGetValue("y", out var yStr) &&
                int.TryParse(xStr?.ToString(), out int x) && 
                int.TryParse(yStr?.ToString(), out int y))
            {
                try
                {
                    Mouse.Position = new System.Drawing.Point(x, y);
                    Mouse.Click();
                    return $"({x}, {y}) koordinatına tıklandı.";
                }
                catch (Exception ex)
                {
                    return $"Koordinata tıklama hatası: {ex.Message}";
                }
            }

            // Yoksa uygulama içinde eleman ara ve tıkla
            string appName = command.Parameters.TryGetValue("application", out var appParam) ? appParam?.ToString() : null;
            
            if (string.IsNullOrEmpty(appName) || !_runningApplications.TryGetValue(appName, out var application))
            {
                return "Tıklama yapılacak uygulama bulunamadı.";
            }

            try
            {
                using (var automation = new UIA3Automation())
                {
                    var window = application.GetMainWindow(automation);
                    
                    // Hedef ismiyle arayüz elemanını bul
                    var condition = new PropertyCondition(automation.PropertyLibrary.Element.Name, command.Target);
                    var element = window.FindFirstDescendant(condition);
                    
                    if (element != null)
                    {
                        element.Click();
                        return $"{command.Target} elemanına tıklandı.";
                    }
                    
                    return $"{command.Target} elemanı bulunamadı.";
                }
            }
            catch (Exception ex)
            {
                return $"Tıklama hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Searches on the web or in an application
        /// </summary>
        private async Task<string> SearchAsync(GeminiCommand command)
        {
            if (string.IsNullOrEmpty(command.Target))
            {
                return "Arama sorgusu belirtilmedi.";
            }

            string searchEngine = command.Parameters.TryGetValue("engine", out var engine) ? engine?.ToString().ToLower() : "google";
            string searchQuery = Uri.EscapeDataString(command.Target);
            
            try
            {
                string url = searchEngine switch
                {
                    "google" => $"https://www.google.com/search?q={searchQuery}",
                    "bing" => $"https://www.bing.com/search?q={searchQuery}",
                    "yahoo" => $"https://search.yahoo.com/search?p={searchQuery}",
                    "duckduckgo" => $"https://duckduckgo.com/?q={searchQuery}",
                    "yandex" => $"https://yandex.com/search/?text={searchQuery}",
                    _ => $"https://www.google.com/search?q={searchQuery}"
                };

                // Opens the search URL in the default web browser
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });

                return $"\"{command.Target}\" için {searchEngine} üzerinde arama yapıldı.";
            }
            catch (Exception ex)
            {
                return $"Arama hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Conducts research on a specific topic
        /// </summary>
        private async Task<string> ResearchAsync(GeminiCommand command)
        {
            if (string.IsNullOrEmpty(command.Target))
            {
                return "Araştırma konusu belirtilmedi.";
            }

            try
            {
                // Hedef dosya belirtilmişse sonuçları kaydet
                string targetFile = command.Parameters.TryGetValue("outputFile", out var file) ? file?.ToString() : null;
                
                // Gemini API kullanarak araştırma yap (aynı API'yi araştırma için de kullanıyoruz)
                var researchRequest = new
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
                                    text = $"Lütfen aşağıdaki konu hakkında detaylı bir araştırma yap ve sonuçları özet olarak paylaş:\n\n{command.Target}"
                                }
                            }
                        }
                    }
                };

                var requestUrl = $"{_geminiService._apiUrl}?key={_geminiService._apiKey}";
                string jsonContent = System.Text.Json.JsonSerializer.Serialize(researchRequest);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                
                var response = await _httpClient.PostAsync(requestUrl, content);
                response.EnsureSuccessStatusCode();

                var responseContent = await response.Content.ReadAsStringAsync();
                var parsedResponse = System.Text.Json.JsonDocument.Parse(responseContent);

                // Gemini API yanıt yapısından metni çıkar
                var text = parsedResponse
                    .RootElement
                    .GetProperty("candidates")[0]
                    .GetProperty("content")
                    .GetProperty("parts")[0]
                    .GetProperty("text")
                    .GetString();

                // Hedef dosya belirtilmişse sonuçları kaydet
                if (!string.IsNullOrEmpty(targetFile))
                {
                    File.WriteAllText(targetFile, text);
                    return $"\"{command.Target}\" konusu araştırıldı ve sonuçlar {targetFile} dosyasına kaydedildi.";
                }

                return text;
            }
            catch (Exception ex)
            {
                return $"Araştırma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Reads a file
        /// </summary>
        private string ReadFile(GeminiCommand command)
        {
            if (string.IsNullOrEmpty(command.Target))
            {
                return "Okunacak dosya belirtilmedi.";
            }

            try
            {
                if (!File.Exists(command.Target))
                {
                    return $"Dosya bulunamadı: {command.Target}";
                }

                // Dosya kodlaması
                string encoding = command.Parameters.TryGetValue("encoding", out var enc) ? enc?.ToString().ToLower() : "utf8";
                Encoding fileEncoding = encoding switch
                {
                    "ascii" => Encoding.ASCII,
                    "unicode" => Encoding.Unicode,
                    "utf7" => Encoding.UTF7,
                    "utf32" => Encoding.UTF32,
                    "utf8bom" => Encoding.UTF8, // BOM ile birlikte
                    _ => new UTF8Encoding(false) // BOM olmadan UTF8
                };

                string fileContent = File.ReadAllText(command.Target, fileEncoding);
                
                // İçerik çok uzunsa kısalt
                int maxLength = 2000;
                if (fileContent.Length > maxLength && !command.Parameters.ContainsKey("full"))
                {
                    fileContent = fileContent.Substring(0, maxLength) + $"\n\n... (devamı var, toplam {fileContent.Length} karakter)";
                }
                
                return fileContent;
            }
            catch (Exception ex)
            {
                return $"Dosya okuma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Writes to a file
        /// </summary>
        private string WriteFile(GeminiCommand command)
        {
            if (string.IsNullOrEmpty(command.Target))
            {
                return "Yazılacak dosya belirtilmedi.";
            }

            if (!command.Parameters.TryGetValue("content", out var contentObj))
            {
                return "Yazılacak içerik belirtilmedi.";
            }
            string content = contentObj?.ToString();

            try
            {
                // Dosya kodlaması
                string encoding = command.Parameters.TryGetValue("encoding", out var enc) ? enc?.ToString().ToLower() : "utf8";
                Encoding fileEncoding = encoding switch
                {
                    "ascii" => Encoding.ASCII,
                    "unicode" => Encoding.Unicode,
                    "utf7" => Encoding.UTF7,
                    "utf32" => Encoding.UTF32,
                    "utf8bom" => Encoding.UTF8, // BOM ile birlikte
                    _ => new UTF8Encoding(false) // BOM olmadan UTF8
                };

                // Dosya içeriğine ekleme yapılacak mı?
                bool append = command.Parameters.TryGetValue("append", out var appendValue) && 
                             (appendValue is bool boolValue ? boolValue : 
                              appendValue?.ToString().ToLower() == "true" || appendValue?.ToString() == "1");
                
                // Dosyaya yaz
                if (append)
                {
                    File.AppendAllText(command.Target, content, fileEncoding);
                    return $"{command.Target} dosyasına içerik eklendi.";
                }
                else
                {
                    File.WriteAllText(command.Target, content, fileEncoding);
                    return $"{command.Target} dosyasına içerik yazıldı.";
                }
            }
            catch (Exception ex)
            {
                return $"Dosya yazma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Runs a script
        /// </summary>
        private string RunScript(GeminiCommand command)
        {
            if (string.IsNullOrEmpty(command.Target))
            {
                return "Çalıştırılacak script belirtilmedi.";
            }

            try
            {
                if (!File.Exists(command.Target))
                {
                    return $"Script dosyası bulunamadı: {command.Target}";
                }

                string extension = Path.GetExtension(command.Target).ToLower();
                var startInfo = new ProcessStartInfo();
                
                // Betik türüne göre çalıştırıcıyı belirle
                switch (extension)
                {
                    case ".bat":
                    case ".cmd":
                        startInfo.FileName = "cmd.exe";
                        startInfo.Arguments = $"/c \"{command.Target}\"";
                        break;
                    case ".ps1":
                        startInfo.FileName = "powershell.exe";
                        startInfo.Arguments = $"-ExecutionPolicy Bypass -File \"{command.Target}\"";
                        break;
                    case ".py":
                        startInfo.FileName = "python";
                        startInfo.Arguments = $"\"{command.Target}\"";
                        break;
                    case ".js":
                        startInfo.FileName = "node";
                        startInfo.Arguments = $"\"{command.Target}\"";
                        break;
                    default:
                        // Doğrudan çalıştırabilir olduğunu varsay
                        startInfo.FileName = command.Target;
                        break;
                }
                
                // Ek argümanlar
                if (command.Parameters.TryGetValue("arguments", out var args))
                {
                    startInfo.Arguments += " " + args?.ToString();
                }
                
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.CreateNoWindow = true;
                
                var process = Process.Start(startInfo);
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();
                
                if (!string.IsNullOrEmpty(error))
                {
                    return $"Script çalıştırıldı, ancak hatalar oluştu:\n{error}";
                }
                
                return $"Script başarıyla çalıştırıldı.\nÇıktı:\n{output}";
            }
            catch (Exception ex)
            {
                return $"Script çalıştırma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Gets system information
        /// </summary>
        private string GetSystemInfo(GeminiCommand command)
        {
            string infoType = command.Target?.ToLower() ?? "general";
            
            var sb = new StringBuilder("Sistem Bilgileri:\n");
            
            try
            {
                switch (infoType)
                {
                    case "os":
                    case "işletim sistemi":
                        sb.AppendLine($"İşletim Sistemi: {Environment.OSVersion}");
                        sb.AppendLine($"64-bit: {Environment.Is64BitOperatingSystem}");
                        sb.AppendLine($"Makine Adı: {Environment.MachineName}");
                        sb.AppendLine($"Sistem Dizini: {Environment.SystemDirectory}");
                        break;
                        
                    case "cpu":
                    case "işlemci":
                        sb.AppendLine($"İşlemci Sayısı: {Environment.ProcessorCount}");
                        // WMI ile daha detaylı bilgi alınabilir
                        break;
                        
                    case "memory":
                    case "bellek":
                    case "ram":
                        sb.AppendLine($"Toplam Fiziksel Bellek: {GetTotalPhysicalMemory()} GB");
                        sb.AppendLine($"Kullanılabilir Bellek: {GetAvailableMemory()} GB");
                        break;
                        
                    case "disk":
                    case "depolama":
                        foreach (var drive in DriveInfo.GetDrives())
                        {
                            if (drive.IsReady)
                            {
                                sb.AppendLine($"Sürücü {drive.Name} - Toplam: {drive.TotalSize / (1024 * 1024 * 1024)} GB, Boş: {drive.AvailableFreeSpace / (1024 * 1024 * 1024)} GB");
                            }
                        }
                        break;
                        
                    case "user":
                    case "kullanıcı":
                        sb.AppendLine($"Kullanıcı Adı: {Environment.UserName}");
                        sb.AppendLine($"Kullanıcı Domain: {Environment.UserDomainName}");
                        break;
                        
                    case "running":
                    case "applications":
                    case "apps":
                        return GetRunningApplications();
                        
                    case "general":
                    case "genel":
                    default:
                        sb.AppendLine($"İşletim Sistemi: {Environment.OSVersion}");
                        sb.AppendLine($"64-bit: {Environment.Is64BitOperatingSystem}");
                        sb.AppendLine($"Makine Adı: {Environment.MachineName}");
                        sb.AppendLine($"İşlemci Sayısı: {Environment.ProcessorCount}");
                        sb.AppendLine($"Toplam Fiziksel Bellek: {GetTotalPhysicalMemory()} GB");
                        sb.AppendLine($"Kullanıcı Adı: {Environment.UserName}");
                        break;
                }
                
                return sb.ToString();
            }
            catch (Exception ex)
            {
                return $"Sistem bilgisi alınamadı: {ex.Message}";
            }
        }
        
        private double GetTotalPhysicalMemory()
        {
            try
            {
                return new Microsoft.VisualBasic.Devices.ComputerInfo().TotalPhysicalMemory / (1024.0 * 1024 * 1024);
            }
            catch
            {
                return 0;
            }
        }
        
        private double GetAvailableMemory()
        {
            try
            {
                return new Microsoft.VisualBasic.Devices.ComputerInfo().AvailablePhysicalMemory / (1024.0 * 1024 * 1024);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Returns a list of running applications
        /// </summary>
        private string GetRunningApplications()
        {
            var sb = new StringBuilder("Çalışan Uygulamalar:\n\n");
            
            // NanAI tarafından kontrol edilen uygulamalar
            if (_runningApplications.Count > 0)
            {
                sb.AppendLine("NanAI tarafından kontrol edilen uygulamalar:");
                sb.AppendLine(string.Join(", ", _runningApplications.Keys));
                sb.AppendLine();
            }
            else
            {
                sb.AppendLine("NanAI tarafından kontrol edilen uygulama yok.");
                sb.AppendLine();
            }
            
            // Tüm çalışan işlemler
            var processes = Process.GetProcesses();
            sb.AppendLine($"Tüm çalışan işlemler ({processes.Length}):");
            
            // LINQ ile filtreleme yerine manuel filtreleme kullanıyoruz
            var filteredProcesses = new List<Process>();
            foreach (var proc in processes)
            {
                if (!string.IsNullOrEmpty(proc.MainWindowTitle))
                {
                    filteredProcesses.Add(proc);
                }
            }
            
            // Gruplama işlemi
            var processGroups = filteredProcesses
                .GroupBy(p => p.ProcessName)
                .OrderBy(g => g.Key);
                
            foreach (var group in processGroups)
            {
                sb.AppendLine($"- {group.Key} ({group.Count()})");
            }
            
            return sb.ToString();
        }
        
        /// <summary>
        /// Controls Windows services (start, stop, status)
        /// </summary>
        private string ControlService(GeminiCommand command)
        {
            try
            {
                string serviceName = command.Target;
                if (string.IsNullOrEmpty(serviceName))
                {
                    return "Kontrol edilecek servis adı belirtilmedi.";
                }
                
                // İşlem türünü belirle (başlat, durdur, yeniden başlat, durum)
                string action = "status"; // varsayılan: durum kontrolü
                if (command.Parameters.TryGetValue("action", out var actionParam))
                {
                    action = actionParam?.ToString().ToLower();
                }
                
                // Eğer tam servis adı bilinmiyorsa, servis adını bulmaya çalış
                string exactServiceName = FindWindowsService(serviceName);
                if (string.IsNullOrEmpty(exactServiceName))
                {
                    return $"'{serviceName}' adında bir Windows servisi bulunamadı. Servis adını kontrol edin veya tam adını belirtin.";
                }
                
                // İşlemi gerçekleştir
                return action switch
                {
                    "start" => StartWindowsService(exactServiceName),
                    "stop" => StopWindowsService(exactServiceName),
                    "restart" => RestartWindowsService(exactServiceName),
                    "status" => GetWindowsServiceStatus(exactServiceName),
                    _ => $"Desteklenmeyen servis işlemi: {action}. Kullanılabilir işlemler: start, stop, restart, status."
                };
            }
            catch (Exception ex)
            {
                return $"Servis kontrolü hatası: {ex.Message}";
            }
        }
        
        /// <summary>
        /// Finds a Windows service by name
        /// </summary>
        private string FindWindowsService(string serviceName)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "sc",
                        Arguments = "query state= all",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                
                // Extract service names
                var serviceNames = new List<string>();
                string[] lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                
                foreach (var line in lines)
                {
                    if (line.Trim().StartsWith("SERVICE_NAME:"))
                    {
                        string name = line.Substring(line.IndexOf(':') + 1).Trim();
                        serviceNames.Add(name);
                    }
                }
                
                // Look for exact match
                string exactMatch = serviceNames.FirstOrDefault(s => s.Equals(serviceName, StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(exactMatch))
                {
                    return exactMatch;
                }
                
                // Look for partial match - using ToLower() instead of StringComparison parameter
                string partialMatch = serviceNames.FirstOrDefault(s => 
                    s.ToLower().Contains(serviceName.ToLower()) || 
                    serviceName.ToLower().Contains(s.ToLower()));
                
                if (!string.IsNullOrEmpty(partialMatch))
                {
                    return partialMatch;
                }
                
                // Look for similarity
                int threshold = 3; // Levenshtein distance threshold
                var similarServices = serviceNames
                    .Select(s => new { Name = s, Distance = LevenshteinDistance(s.ToLower(), serviceName.ToLower()) })
                    .Where(item => item.Distance <= threshold)
                    .OrderBy(item => item.Distance)
                    .ToList();
                    
                if (similarServices.Count > 0)
                {
                    return similarServices[0].Name;
                }
                
                // Not found
                return "";
            }
            catch
            {
                // Return empty on error, will be handled in the main method
                return "";
            }
        }
        
        /// <summary>
        /// Starts a Windows service
        /// </summary>
        private string StartWindowsService(string serviceName)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "sc",
                        Arguments = $"start \"{serviceName}\"",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                
                // Hata durumunu kontrol et
                if (output.Contains("FAILED") || output.Contains("ERROR"))
                {
                    return $"'{serviceName}' servisi başlatılamadı. Hata: {output.Trim()}";
                }
                
                // Net start komutunu da dene (daha güvenilir olabilir)
                try
                {
                    var netStartProcess = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = "net",
                            Arguments = $"start \"{serviceName}\"",
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };
                    
                    netStartProcess.Start();
                    netStartProcess.WaitForExit();
                }
                catch
                {
                    // Net start hatası yoksay, SC komutu çalıştı bile
                }
                
                return $"'{serviceName}' servisi başlatıldı.";
            }
            catch (Exception ex)
            {
                return $"Servis başlatma hatası: {ex.Message}";
            }
        }
        
        /// <summary>
        /// Stops a Windows service
        /// </summary>
        private string StopWindowsService(string serviceName)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "sc",
                        Arguments = $"stop \"{serviceName}\"",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                
                // Hata durumunu kontrol et
                if (output.Contains("FAILED") || output.Contains("ERROR"))
                {
                    return $"'{serviceName}' servisi durdurulamadı. Hata: {output.Trim()}";
                }
                
                // Net stop komutunu da dene
                try
                {
                    var netStopProcess = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = "net",
                            Arguments = $"stop \"{serviceName}\"",
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };
                    
                    netStopProcess.Start();
                    netStopProcess.WaitForExit();
                }
                catch
                {
                    // Net stop hatası yoksay, SC komutu çalıştı bile
                }
                
                return $"'{serviceName}' servisi durduruldu.";
            }
            catch (Exception ex)
            {
                return $"Servis durdurma hatası: {ex.Message}";
            }
        }
        
        /// <summary>
        /// Restarts a Windows service
        /// </summary>
        private string RestartWindowsService(string serviceName)
        {
            string stopResult = StopWindowsService(serviceName);
            
            if (stopResult.Contains("durduruldu"))
            {
                // Servisin durması için biraz bekle
                System.Threading.Thread.Sleep(3000);
                
                string startResult = StartWindowsService(serviceName);
                return $"'{serviceName}' servisi yeniden başlatıldı.";
            }
            else
            {
                return stopResult; // Durdurma hatası olduysa onu göster
            }
        }
        
        /// <summary>
        /// Shows the status of a Windows service
        /// </summary>
        private string GetWindowsServiceStatus(string serviceName)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "sc",
                        Arguments = $"query \"{serviceName}\"",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                
                if (output.Contains("FAILED") || output.Contains("ERROR") || output.Contains("does not exist"))
                {
                    return $"'{serviceName}' servisi bulunamadı.";
                }
                
                // Durumu çıkar
                string state = "BİLİNMİYOR";
                string[] lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                
                foreach (var line in lines)
                {
                    if (line.Trim().StartsWith("STATE"))
                    {
                        if (line.Contains("RUNNING"))
                            state = "ÇALIŞIYOR";
                        else if (line.Contains("STOPPED"))
                            state = "DURDURULDU";
                        else if (line.Contains("STARTING"))
                            state = "BAŞLATILIYOR";
                        else if (line.Contains("STOPPING"))
                            state = "DURDURULUYOR";
                        else if (line.Contains("PAUSED"))
                            state = "DURAKLATILDI";
                        
                        break;
                    }
                }
                
                StringBuilder result = new StringBuilder();
                result.AppendLine($"'{serviceName}' servisi durumu: {state}");
                
                // Ek bilgileri ekle
                foreach (var line in lines)
                {
                    string trimmedLine = line.Trim();
                    if (trimmedLine.StartsWith("DISPLAY_NAME"))
                    {
                        string displayName = trimmedLine.Substring(trimmedLine.IndexOf(':') + 1).Trim();
                        result.AppendLine($"Görünen Ad: {displayName}");
                    }
                    else if (trimmedLine.StartsWith("WIN32_EXIT_CODE") && !trimmedLine.Contains("0"))
                    {
                        string exitCode = trimmedLine.Substring(trimmedLine.IndexOf(':') + 1).Trim();
                        result.AppendLine($"Çıkış Kodu: {exitCode}");
                    }
                    else if (trimmedLine.StartsWith("PID"))
                    {
                        string pid = trimmedLine.Substring(trimmedLine.IndexOf(':') + 1).Trim();
                        result.AppendLine($"İşlem ID: {pid}");
                    }
                }
                
                return result.ToString();
            }
            catch (Exception ex)
            {
                return $"Servis durumu sorgulama hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Handles custom commands that do not fit the standard format
        /// </summary>
        private async Task<string> ProcessCustomCommandAsync(GeminiCommand command)
        {
            // For the basic application commands, try to execute them directly
            string lowerPrompt = command.Target.ToLower().Trim();
            
            // Handle "open" command directly
            if (lowerPrompt.StartsWith("open "))
            {
                string appName = lowerPrompt.Substring(5).Trim();
                if (!string.IsNullOrEmpty(appName))
                {
                    var launchCommand = new GeminiCommand
                    {
                        CommandType = "launch",
                        Target = appName,
                        Parameters = new Dictionary<string, object>()
                    };
                    return LaunchApplication(launchCommand);
                }
            }
            
            // Handle "write" command directly
            if (lowerPrompt.StartsWith("write "))
            {
                string text = command.Target.Substring(6).Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    var typeCommand = new GeminiCommand
                    {
                        CommandType = "type",
                        Target = "",
                        Parameters = new Dictionary<string, object>
                        {
                            { "text", text }
                        }
                    };
                    return TypeText(typeCommand);
                }
            }
            
            // If it's a system info request
            if (lowerPrompt.Contains("system") && lowerPrompt.Contains("info"))
            {
                var sysInfoCommand = new GeminiCommand
                {
                    CommandType = "systeminfo",
                    Target = "general",
                    Parameters = new Dictionary<string, object>()
                };
                return GetSystemInfo(sysInfoCommand);
            }
            
            // Use the Gemini API to generate a response for custom commands
            try
            {
                // Generate a more helpful message when API fails
                if (lowerPrompt.Contains("custom command not processed"))
                {
                    return "I'm sorry, I couldn't process that command. Please try with a more specific command such as 'open notepad', 'write hello', or 'show system info'.";
                }
                
                var customRequest = new
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
                                    text = command.Target
                                }
                            }
                        }
                    }
                };

                var requestUrl = $"{_geminiService._apiUrl}?key={_geminiService._apiKey}";
                string jsonContent = System.Text.Json.JsonSerializer.Serialize(customRequest);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                
                // Wrap in try-catch to handle API errors gracefully
                try
                {
                    var response = await _httpClient.PostAsync(requestUrl, content);
                    
                    if (!response.IsSuccessStatusCode)
                    {
                        return "I couldn't process that command. Please try with a specific action like 'open [application]' or 'write [text]'.";
                    }
                    
                    var responseContent = await response.Content.ReadAsStringAsync();
                    var parsedResponse = System.Text.Json.JsonDocument.Parse(responseContent);

                    // Gemini API yanıt yapısından metni çıkar
                    var text = parsedResponse
                        .RootElement
                        .GetProperty("candidates")[0]
                        .GetProperty("content")
                        .GetProperty("parts")[0]
                        .GetProperty("text")
                        .GetString();

                    return text;
                }
                catch
                {
                    // Fall back to direct command execution
                    return TryDirectCommand(command.Target);
                }
            }
            catch (Exception ex)
            {
                return $"Custom command not processed: {ex.Message}";
            }
        }
        
        /// <summary>
        /// Attempts to execute a command directly without API
        /// </summary>
        private string TryDirectCommand(string command)
        {
            string lowerCommand = command.ToLower().Trim();
            
            // Try to identify and execute common commands
            if (lowerCommand == "notepad" || lowerCommand == "open notepad")
            {
                var launchCommand = new GeminiCommand
                {
                    CommandType = "launch",
                    Target = "notepad",
                    Parameters = new Dictionary<string, object>()
                };
                return LaunchApplication(launchCommand);
            }
            
            if (lowerCommand == "calculator" || lowerCommand == "open calculator")
            {
                var launchCommand = new GeminiCommand
                {
                    CommandType = "launch",
                    Target = "calculator",
                    Parameters = new Dictionary<string, object>()
                };
                return LaunchApplication(launchCommand);
            }
            
            if (lowerCommand.Contains("system") && (lowerCommand.Contains("info") || lowerCommand.Contains("information")))
            {
                var sysInfoCommand = new GeminiCommand
                {
                    CommandType = "systeminfo",
                    Target = "general",
                    Parameters = new Dictionary<string, object>()
                };
                return GetSystemInfo(sysInfoCommand);
            }
            
            // If no direct match, attempt to parse as a launch command
            string[] words = lowerCommand.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length > 0)
            {
                var launchCommand = new GeminiCommand
                {
                    CommandType = "launch",
                    Target = words[0],
                    Parameters = new Dictionary<string, object>()
                };
                return LaunchApplication(launchCommand);
            }
            
            return "I couldn't understand that command. Please try with a specific action like 'open [application]' or 'write [text]'.";
        }

        /// <summary>
        /// Handles natural language service commands
        /// </summary>
        private async Task<string> ControlServiceAsync(string userInput)
        {
            try
            {
                // Try to extract service name and command
                string normalizedInput = userInput.ToLower();
                string serviceName = "";
                string action = "";
                
                // Determine the action type
                if (normalizedInput.Contains("start") || normalizedInput.Contains("run"))
                {
                    action = "start";
                }
                else if (normalizedInput.Contains("stop") || normalizedInput.Contains("kill") ||
                        normalizedInput.Contains("terminate"))
                {
                    action = "stop";
                }
                else if (normalizedInput.Contains("restart") || normalizedInput.Contains("reset") ||
                        normalizedInput.Contains("refresh"))
                {
                    action = "restart";
                }
                else
                {
                    action = "status"; // Default to status check
                }
                
                // Use some patterns to find the service name
                string[] patterns = new[] {
                    "service(?:s)? (.*?) (?:start|run|stop|restart|check|status)",
                    "(.*?) service(?:s)? (?:start|run|stop|restart|check|status)",
                    "(?:start|run|stop|restart|check|status) (.*?) service(?:s)?"
                };
                
                foreach (var pattern in patterns)
                {
                    var match = System.Text.RegularExpressions.Regex.Match(normalizedInput, pattern);
                    if (match.Success && match.Groups.Count > 1)
                    {
                        serviceName = match.Groups[1].Value.Trim();
                        break;
                    }
                }
                
                // Another strategy: take the first word after "service"
                if (string.IsNullOrEmpty(serviceName))
                {
                    string[] keywords = new[] { "service" };
                    foreach (var keyword in keywords)
                    {
                        int index = normalizedInput.IndexOf(keyword);
                        if (index >= 0)
                        {
                            string afterKeyword = normalizedInput.Substring(index + keyword.Length).Trim();
                            if (!string.IsNullOrEmpty(afterKeyword))
                            {
                                // Take the first word
                                string[] parts = afterKeyword.Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                                if (parts.Length > 0)
                                {
                                    serviceName = parts[0].Trim();
                                    break;
                                }
                            }
                        }
                    }
                }
                
                // If still not found: remove action words and guess from remaining text
                if (string.IsNullOrEmpty(serviceName))
                {
                    string[] actionWords = new[] { 
                        "start", "run", 
                        "stop", "kill", "terminate",
                        "restart", "reset", "refresh",
                        "check", "status",
                        "service"
                    };
                    
                    string cleaned = normalizedInput;
                    foreach (var word in actionWords)
                    {
                        cleaned = cleaned.Replace(word, " ");
                    }
                    
                    // Clean the remaining text and consider the first word as service name
                    cleaned = cleaned.Trim();
                    if (!string.IsNullOrEmpty(cleaned))
                    {
                        string[] parts = cleaned.Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length > 0)
                        {
                            serviceName = parts[0].Trim();
                        }
                    }
                }
                
                // If still not found, try to use Gemini API
                if (string.IsNullOrEmpty(serviceName))
                {
                    // Create a specific request
                    string prompt = $"Try to extract a Windows service name and the desired action (start, stop, check status) from this text. Return only the service name and action in JSON format. Example: {{\"serviceName\": \"spooler\", \"action\": \"start\"}}. Here's the text: \"{userInput}\"";
                    
                    var response = await _geminiService.SendRawPromptAsync(prompt);
                    try
                    {
                        var parsedResponse = System.Text.Json.JsonDocument.Parse(response);
                        if (parsedResponse.RootElement.TryGetProperty("serviceName", out var serviceNameElement))
                        {
                            serviceName = serviceNameElement.GetString();
                        }
                        
                        if (parsedResponse.RootElement.TryGetProperty("action", out var actionElement))
                        {
                            action = actionElement.GetString();
                        }
                    }
                    catch
                    {
                        // JSON parsing error, don't use the response directly
                    }
                }
                
                if (string.IsNullOrEmpty(serviceName))
                {
                    return "Could not determine the service name to control. Please specify the service name explicitly.";
                }
                
                // Create and process GeminiCommand
                var command = new GeminiCommand
                {
                    CommandType = "ServiceControl",
                    Target = serviceName,
                    Parameters = new Dictionary<string, object>
                    {
                        { "action", action }
                    }
                };
                
                return ControlService(command);
            }
            catch (Exception ex)
            {
                return $"Service control error: {ex.Message}";
            }
        }
    }
} 