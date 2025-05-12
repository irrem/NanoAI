using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoAI.Gemini;
using System.Windows.Forms;

namespace NanoAI.Core.NLU.Handlers
{
    /// <summary>
    /// Handles composite commands that contain multiple sequential instructions
    /// </summary>
    public class CompositeCommandHandler : ICommandHandler
    {
        public string CommandType => "composite";
        
        // Reference to the parent service to process sub-commands
        public EnhancedGeminiService ParentService { get; set; }
        
        // Delimiters used to split a composite command into individual commands
        private readonly string[] _delimiters = new string[] 
        { 
            " then ", 
            " and then ", 
            " followed by ", 
            ", then ", 
            " after that ", 
            " next ", 
            ", next ", 
            " and next ", 
            ", and then ",
            " and "
        };
        
        public CompositeCommandHandler(EnhancedGeminiService service)
        {
            ParentService = service ?? throw new ArgumentNullException(nameof(service));
        }
        
        public bool CanHandle(GeminiCommand command)
        {
            if (command == null || string.IsNullOrEmpty(command.Target))
                return false;
                
            // Check if the command contains any of the sequence delimiters
            foreach (var delimiter in _delimiters)
            {
                if (command.Target.IndexOf(delimiter, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            
            return false;
        }
        
        public async Task<CommandResult> ExecuteAsync(GeminiCommand command)
        {
            if (string.IsNullOrEmpty(command.Target))
            {
                return new CommandResult
                {
                    Success = false,
                    Message = "No command sequence provided",
                    Suggestions = new List<string> { "Please provide a sequence of commands to execute" }
                };
            }
            
            var commands = SplitIntoCommandSequence(command.Target);
            
            if (commands.Count == 0)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = "Could not parse command sequence",
                    Suggestions = new List<string> { "Please check your command syntax" }
                };
            }
            
            if (commands.Count == 1)
            {
                // If there's only one command, just process it directly
                return await ParentService.ProcessCommandAsync(commands[0]);
            }
            
            var results = new List<string>();
            var allErrors = new List<string>();
            var anySuccess = false;
            
            // Check for special patterns in the entire command string
            string fullCommand = command.Target.ToLower();
            
            // Handle "open X and write Y" pattern even when there are more than 2 commands
            if (commands.Count >= 2)
            {
                // Find the last command that contains "write" or "type"
                int writeCommandIndex = -1;
                for (int i = commands.Count - 1; i >= 0; i--)
                {
                    string lowerCmd = commands[i].ToLower().Trim();
                    if (lowerCmd.StartsWith("write ") || lowerCmd.StartsWith("type "))
                    {
                        writeCommandIndex = i;
                        break;
                    }
                }
                
                // Find the last "open", "launch", or "start" command before the write command
                int openCommandIndex = -1;
                if (writeCommandIndex > 0)
                {
                    for (int i = writeCommandIndex - 1; i >= 0; i--)
                    {
                        string lowerCmd = commands[i].ToLower().Trim();
                        if (lowerCmd.StartsWith("open ") || lowerCmd.StartsWith("launch ") || lowerCmd.StartsWith("start "))
                        {
                            openCommandIndex = i;
                            break;
                        }
                    }
                }
                
                // If we found both an open and write command pattern
                if (openCommandIndex >= 0 && writeCommandIndex > openCommandIndex)
                {
                    string appName = GetAppNameFromCommand(commands[openCommandIndex]);
                    string text = GetTextFromCommand(commands[writeCommandIndex]);
                    
                    Console.WriteLine($"Detected open-and-write pattern: App={appName}, Text={text}");
                    
                    if (!string.IsNullOrEmpty(appName) && !string.IsNullOrEmpty(text))
                    {
                        // Process all commands before the open command
                        for (int i = 0; i < openCommandIndex; i++)
                        {
                            try
                            {
                                var result = await ParentService.ProcessCommandAsync(commands[i]);
                                if (result.Success)
                                {
                                    results.Add($"Step {i+1}: {result.Message}");
                                    anySuccess = true;
                                }
                                else
                                {
                                    results.Add($"Step {i+1} failed: {result.Message}");
                                    if (result.Suggestions != null)
                                    {
                                        allErrors.AddRange(result.Suggestions);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                results.Add($"Step {i+1} error: {ex.Message}");
                                allErrors.Add(ex.Message);
                            }
                            
                            // Give a short delay between commands
                            await Task.Delay(500);
                        }
                        
                        // Launch the application
                        try
                        {
                            var launchCommand = new GeminiCommand
                            {
                                CommandType = "launch",
                                Target = appName,
                                Parameters = new Dictionary<string, object>()
                            };
                            
                            var launchResult = await ParentService.ExecuteCommandDirectlyAsync(launchCommand);
                            results.Add($"Step {openCommandIndex+1}: {launchResult.Message}");
                            
                            if (launchResult.Success)
                            {
                                // Process any commands between the open and write command
                                for (int i = openCommandIndex + 1; i < writeCommandIndex; i++)
                                {
                                    try
                                    {
                                        var result = await ParentService.ProcessCommandAsync(commands[i]);
                                        if (result.Success)
                                        {
                                            results.Add($"Step {i+1}: {result.Message}");
                                            anySuccess = true;
                                        }
                                        else
                                        {
                                            results.Add($"Step {i+1} failed: {result.Message}");
                                            if (result.Suggestions != null)
                                            {
                                                allErrors.AddRange(result.Suggestions);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        results.Add($"Step {i+1} error: {ex.Message}");
                                        allErrors.Add(ex.Message);
                                    }
                                    
                                    // Give a short delay between commands
                                    await Task.Delay(500);
                                }
                                
                                // Give the app time to initialize
                                await Task.Delay(2000);
                                
                                // Then type the text
                                try
                                {
                                    // Use simpler SendKeys approach for typing text
                                    SendKeys.SendWait(text);
                                    results.Add($"Step {writeCommandIndex+1}: Typed '{text}' into {appName}");
                                    anySuccess = true;
                                    
                                    // Process any remaining commands
                                    for (int i = writeCommandIndex + 1; i < commands.Count; i++)
                                    {
                                        try
                                        {
                                            var result = await ParentService.ProcessCommandAsync(commands[i]);
                                            if (result.Success)
                                            {
                                                results.Add($"Step {i+1}: {result.Message}");
                                                anySuccess = true;
                                            }
                                            else
                                            {
                                                results.Add($"Step {i+1} failed: {result.Message}");
                                                if (result.Suggestions != null)
                                                {
                                                    allErrors.AddRange(result.Suggestions);
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            results.Add($"Step {i+1} error: {ex.Message}");
                                            allErrors.Add(ex.Message);
                                        }
                                        
                                        // Give a short delay between commands
                                        await Task.Delay(500);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string errorMsg = $"Error typing text: {ex.Message}";
                                    results.Add($"Step {writeCommandIndex+1} failed: {errorMsg}");
                                    allErrors.Add(errorMsg);
                                }
                            }
                            else
                            {
                                allErrors.Add(launchResult.Message);
                            }
                            
                            // Build the final result
                            var innerResultBuilder = new StringBuilder();
                            foreach (var result in results)
                            {
                                innerResultBuilder.AppendLine(result);
                            }
                            
                            return new CommandResult
                            {
                                Success = anySuccess,
                                Message = innerResultBuilder.ToString().Trim(),
                                Suggestions = allErrors.Count > 0 ? allErrors : null
                            };
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing open-and-write pattern: {ex.Message}");
                        }
                    }
                }
            }
            
            // If we didn't handle it as a special pattern, process commands sequentially
            for (int i = 0; i < commands.Count; i++)
            {
                string cmdText = commands[i];
                
                try
                {
                    var result = await ParentService.ProcessCommandAsync(cmdText);
                    
                    if (result.Success)
                    {
                        results.Add($"Step {i+1}: {result.Message}");
                        anySuccess = true;
                    }
                    else
                    {
                        results.Add($"Step {i+1} failed: {result.Message}");
                        if (result.Suggestions != null)
                        {
                            allErrors.AddRange(result.Suggestions);
                        }
                    }
                }
                catch (Exception ex)
                {
                    results.Add($"Step {i+1} error: {ex.Message}");
                    allErrors.Add(ex.Message);
                }
            }
            
            // Build the final result
            var resultBuilder = new StringBuilder();
            foreach (var result in results)
            {
                resultBuilder.AppendLine(result);
            }
            
            return new CommandResult
            {
                Success = anySuccess,
                Message = resultBuilder.ToString().Trim(),
                Suggestions = allErrors.Count > 0 ? allErrors : null
            };
        }
        
        /// <summary>
        /// Extracts the application name from a launch command
        /// </summary>
        private string GetAppNameFromCommand(string command)
        {
            string lowerCommand = command.ToLower().Trim();
            string appName = "";
            
            if (lowerCommand.StartsWith("open "))
                appName = command.Substring(5).Trim();
            else if (lowerCommand.StartsWith("launch "))
                appName = command.Substring(7).Trim();
            else if (lowerCommand.StartsWith("start "))
                appName = command.Substring(6).Trim();
            else if (lowerCommand.StartsWith("run "))
                appName = command.Substring(4).Trim();
                
            // Clean up common suffixes
            string[] suffixesToRemove = {
                " app", " application", " program"
            };
                
            foreach (var suffix in suffixesToRemove)
            {
                if (appName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    appName = appName.Substring(0, appName.Length - suffix.Length).Trim();
                }
            }
            
            return appName;
        }
        
        /// <summary>
        /// Extracts the text to write from a command
        /// </summary>
        private string GetTextFromCommand(string command)
        {
            string lowerCommand = command.ToLower().Trim();
            string text = "";
            
            // Use the original command for extracting the text to preserve case and special characters
            if (lowerCommand.StartsWith("write "))
                text = command.Substring(6).Trim();
            else if (lowerCommand.StartsWith("type "))
                text = command.Substring(5).Trim();
            else if (lowerCommand.StartsWith("input "))
                text = command.Substring(6).Trim();
            else if (lowerCommand.StartsWith("enter "))
                text = command.Substring(6).Trim();
                
            // Clean up common suffixes (case-insensitive but preserve original text)
            string[] suffixesToRemove = {
                " in it", " inside", " in notepad", " in the notepad", " in document",
                " inside it", " to it", " there", " in app", " in application", " in window"
            };
                
            foreach (var suffix in suffixesToRemove)
            {
                if (text.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    text = text.Substring(0, text.Length - suffix.Length).Trim();
                    break; // Only remove one suffix to prevent over-pruning
                }
            }
            
            return text;
        }
        
        private List<string> SplitIntoCommandSequence(string input)
        {
            var commands = new List<string>();
            string remainingText = input;
            
            while (!string.IsNullOrEmpty(remainingText))
            {
                // Try to find a delimiter
                int earliestPos = remainingText.Length;
                string foundDelimiter = null;
                
                foreach (var delimiter in _delimiters)
                {
                    int pos = remainingText.IndexOf(delimiter, StringComparison.OrdinalIgnoreCase);
                    if (pos >= 0 && pos < earliestPos)
                    {
                        earliestPos = pos;
                        foundDelimiter = delimiter;
                    }
                }
                
                if (foundDelimiter != null)
                {
                    // We found a delimiter, add the command before it
                    string command = remainingText.Substring(0, earliestPos).Trim();
                    if (!string.IsNullOrEmpty(command))
                    {
                        commands.Add(command);
                    }
                    
                    // Update the remaining text
                    remainingText = remainingText.Substring(earliestPos + foundDelimiter.Length).Trim();
                }
                else
                {
                    // No more delimiters, add the remaining text as the last command
                    if (!string.IsNullOrEmpty(remainingText))
                    {
                        commands.Add(remainingText.Trim());
                    }
                    break;
                }
            }
            
            return commands;
        }
    }
} 