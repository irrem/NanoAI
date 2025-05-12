using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NanoAI.Gemini;

namespace NanoAI.Core.NLU.Handlers
{
    /// <summary>
    /// Handles commands to run various types of projects and scripts
    /// </summary>
    public class ProjectCommandHandler : ICommandHandler
    {
        private readonly Dictionary<string, string> _projectLaunchers;

        public string CommandType => "project";

        public ProjectCommandHandler()
        {
            _projectLaunchers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Script files
                { ".py", "python" },
                { ".js", "node" },
                { ".html", "start" },  // Opens in default browser
                
                // .NET projects
                { ".sln", "devenv" },
                { ".csproj", "dotnet run --project" },
                { ".fsproj", "dotnet run --project" },
                { ".vbproj", "dotnet run --project" },
                
                // Java
                { ".java", "javac" },
                { ".jar", "java -jar" },
                
                // PHP
                { ".php", "php" },
                
                // Jupyter notebooks
                { ".ipynb", "jupyter notebook" }
            };
        }

        public bool CanHandle(GeminiCommand command)
        {
            return command.CommandType.Equals("project", StringComparison.OrdinalIgnoreCase);
        }

        public async Task<CommandResult> ExecuteAsync(GeminiCommand command)
        {
            if (command == null)
            {
                throw new ArgumentNullException(nameof(command));
            }

            string projectName = command.Target;
            bool runAsAdmin = command.Parameters != null && 
                              command.Parameters.ContainsKey("runAsAdmin") && 
                              command.Parameters["runAsAdmin"].ToString().ToLower() == "true";

            if (string.IsNullOrEmpty(projectName))
            {
                return new CommandResult
                {
                    Success = false,
                    Message = "Project name is missing",
                    Suggestions = new List<string> {
                        "Please specify the project to run.",
                        "Examples:",
                        "  run my python script",
                        "  open my solution file"
                    }
                };
            }

            // Try to find the project file
            string projectPath = FindProject(projectName);
            
            if (projectPath == null)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Could not find project or script '{projectName}'",
                    Suggestions = new List<string> {
                        "Make sure the file exists",
                        "Try providing the full path"
                    }
                };
            }

            // Get file extension
            string extension = Path.GetExtension(projectPath).ToLower();
            
            // Try to find a launcher for this file type
            if (!_projectLaunchers.TryGetValue(extension, out string launcher))
            {
                // Try to launch directly if no launcher is found
                return await LaunchFile(projectPath, runAsAdmin);
            }
            
            // Launch with the appropriate launcher
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo();
                
                if (launcher == "start")
                {
                    // Special case for HTML files - open with default browser
                    psi.FileName = projectPath;
                    psi.UseShellExecute = true;
                }
                else
                {
                    psi.FileName = "cmd.exe";
                    psi.Arguments = $"/c {launcher} \"{projectPath}\"";
                    psi.UseShellExecute = runAsAdmin;
                    
                    if (runAsAdmin)
                    {
                        psi.Verb = "runas";
                    }
                }
                
                Process.Start(psi);
                Console.WriteLine($"Launched project: {projectPath}");
                return new CommandResult
                {
                    Success = true,
                    Message = $"Launched project '{Path.GetFileName(projectPath)}'"
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to launch project '{projectPath}': {ex.Message}");
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to launch project '{projectName}': {ex.Message}",
                    Suggestions = new List<string> {
                        "Try running as administrator",
                        "Check if all required tools are installed"
                    }
                };
            }
        }

        private async Task<CommandResult> LaunchFile(string filePath, bool asAdmin)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };
                
                if (asAdmin)
                {
                    psi.Verb = "runas";
                }

                Process.Start(psi);
                Console.WriteLine($"Launched file: {filePath}");
                return new CommandResult
                {
                    Success = true,
                    Message = $"Launched file: {Path.GetFileName(filePath)}"
                };
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

        private string FindProject(string projectName)
        {
            // Check if it's an absolute path
            if (Path.IsPathRooted(projectName) && File.Exists(projectName))
            {
                return projectName;
            }
            
            // Common directories to search
            var searchDirs = new List<string>
            {
                // Current directory
                Directory.GetCurrentDirectory(),
                
                // Common locations
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                
                // User profile and source directories
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "source"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Source"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Projects"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "projects")
            };
            
            // First try exact file name
            foreach (var dir in searchDirs)
            {
                if (Directory.Exists(dir))
                {
                    // Search in this directory and subdirectories
                    try
                    {
                        var files = Directory.GetFiles(dir, projectName, SearchOption.AllDirectories);
                        if (files.Length > 0)
                        {
                            return files[0];
                        }
                    }
                    catch
                    {
                        // Skip if access denied
                    }
                }
            }
            
            // Try with common extensions
            foreach (var extension in _projectLaunchers.Keys)
            {
                string nameWithExt = projectName;
                if (!projectName.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                {
                    nameWithExt = projectName + extension;
                }
                
                foreach (var dir in searchDirs)
                {
                    if (Directory.Exists(dir))
                    {
                        try
                        {
                            var files = Directory.GetFiles(dir, nameWithExt, SearchOption.AllDirectories);
                            if (files.Length > 0)
                            {
                                return files[0];
                            }
                        }
                        catch
                        {
                            // Skip if access denied
                        }
                    }
                }
            }
            
            // Try finding by partial match
            foreach (var dir in searchDirs)
            {
                if (Directory.Exists(dir))
                {
                    try
                    {
                        var allFiles = Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories)
                            .Where(f => _projectLaunchers.Keys.Contains(Path.GetExtension(f).ToLower()));
                            
                        foreach (var file in allFiles)
                        {
                            string fileName = Path.GetFileNameWithoutExtension(file);
                            
                            // Check if the file name contains our search string
                            if (fileName.IndexOf(projectName, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return file;
                            }
                        }
                    }
                    catch
                    {
                        // Skip if access denied
                    }
                }
            }
            
            return null;
        }
    }
} 