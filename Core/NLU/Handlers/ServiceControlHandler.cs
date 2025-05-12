using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;
using NanoAI.Gemini;

namespace NanoAI.Core.NLU.Handlers
{
    /// <summary>
    /// Handles commands to control Windows services
    /// </summary>
    public class ServiceControlHandler : ICommandHandler
    {
        public string CommandType => "servicecontrol";

        public bool CanHandle(GeminiCommand command)
        {
            return command.CommandType.Equals("servicecontrol", StringComparison.OrdinalIgnoreCase) ||
                   command.CommandType.Equals("service", StringComparison.OrdinalIgnoreCase);
        }

        public async Task<CommandResult> ExecuteAsync(GeminiCommand command)
        {
            if (command == null)
            {
                throw new ArgumentNullException(nameof(command));
            }

            string serviceName = command.Target;
            string action = command.GetParameterValue("action")?.ToLower() ?? "status";

            if (string.IsNullOrEmpty(serviceName))
            {
                return new CommandResult
                {
                    Success = false,
                    Message = "Service name is missing",
                    Suggestions = new List<string> {
                        "Please specify the service name",
                        "Example: service target=wuauserv action=status"
                    }
                };
            }

            try
            {
                ServiceController service = GetServiceByName(serviceName);
                
                if (service == null)
                {
                    return new CommandResult
                    {
                        Success = false,
                        Message = $"Service '{serviceName}' not found",
                        Suggestions = new List<string> {
                            "Check if the service name is correct",
                            "Try using the exact service name"
                        }
                    };
                }

                switch (action)
                {
                    case "start":
                        return StartService(service);
                    case "stop":
                        return StopService(service);
                    case "restart":
                        return RestartService(service);
                    case "status":
                    default:
                        return GetServiceStatus(service);
                }
            }
            catch (Exception ex)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Error controlling service '{serviceName}': {ex.Message}"
                };
            }
        }

        private ServiceController GetServiceByName(string serviceName)
        {
            // Try exact match first
            var services = ServiceController.GetServices();
            var service = services.FirstOrDefault(s => s.ServiceName.Equals(serviceName, StringComparison.OrdinalIgnoreCase));
            
            // If not found, try contains match
            if (service == null)
            {
                service = services.FirstOrDefault(s => s.ServiceName.ToLower().Contains(serviceName.ToLower()) ||
                                                  s.DisplayName.ToLower().Contains(serviceName.ToLower()));
            }
            
            return service;
        }

        private CommandResult StartService(ServiceController service)
        {
            try
            {
                if (service.Status == ServiceControllerStatus.Running)
                {
                    return new CommandResult
                    {
                        Success = true,
                        Message = $"Service '{service.DisplayName}' is already running"
                    };
                }

                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                
                return new CommandResult
                {
                    Success = true,
                    Message = $"Service '{service.DisplayName}' started successfully"
                };
            }
            catch (Exception ex)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to start service '{service.DisplayName}': {ex.Message}",
                    Suggestions = new List<string> {
                        "Make sure you have administrator privileges",
                        "Check if the service can be started manually"
                    }
                };
            }
        }

        private CommandResult StopService(ServiceController service)
        {
            try
            {
                if (service.Status == ServiceControllerStatus.Stopped)
                {
                    return new CommandResult
                    {
                        Success = true,
                        Message = $"Service '{service.DisplayName}' is already stopped"
                    };
                }

                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                
                return new CommandResult
                {
                    Success = true,
                    Message = $"Service '{service.DisplayName}' stopped successfully"
                };
            }
            catch (Exception ex)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to stop service '{service.DisplayName}': {ex.Message}",
                    Suggestions = new List<string> {
                        "Make sure you have administrator privileges",
                        "Check if the service can be stopped manually"
                    }
                };
            }
        }

        private CommandResult RestartService(ServiceController service)
        {
            try
            {
                if (service.Status == ServiceControllerStatus.Running)
                {
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                }

                service.Start();
                service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                
                return new CommandResult
                {
                    Success = true,
                    Message = $"Service '{service.DisplayName}' restarted successfully"
                };
            }
            catch (Exception ex)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to restart service '{service.DisplayName}': {ex.Message}",
                    Suggestions = new List<string> {
                        "Make sure you have administrator privileges",
                        "Check if the service can be restarted manually"
                    }
                };
            }
        }

        private CommandResult GetServiceStatus(ServiceController service)
        {
            try
            {
                service.Refresh();
                
                string statusText = service.Status.ToString();
                bool isRunning = service.Status == ServiceControllerStatus.Running;
                
                string message = $"Service '{service.DisplayName}' status: {statusText}";
                
                if (isRunning)
                {
                    message += $"\nStartup type: {GetServiceStartupType(service.ServiceName)}";
                }
                
                return new CommandResult
                {
                    Success = true,
                    Message = message
                };
            }
            catch (Exception ex)
            {
                return new CommandResult
                {
                    Success = false,
                    Message = $"Failed to get status for service '{service.DisplayName}': {ex.Message}"
                };
            }
        }

        private string GetServiceStartupType(string serviceName)
        {
            try
            {
                // Use WMI to get startup type
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "sc",
                        Arguments = $"qc \"{serviceName}\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                
                // Parse output for start type
                string startType = "Unknown";
                
                if (output.Contains("START_TYPE") || output.Contains("START TYPE"))
                {
                    if (output.Contains("AUTO_START") || output.Contains("2  AUTO_START"))
                        startType = "Automatic";
                    else if (output.Contains("DEMAND_START") || output.Contains("3  DEMAND_START"))
                        startType = "Manual";
                    else if (output.Contains("DISABLED") || output.Contains("4  DISABLED"))
                        startType = "Disabled";
                    else if (output.Contains("BOOT_START") || output.Contains("0  BOOT_START"))
                        startType = "Boot";
                    else if (output.Contains("SYSTEM_START") || output.Contains("1  SYSTEM_START"))
                        startType = "System";
                }
                
                return startType;
            }
            catch
            {
                return "Unknown";
            }
        }
    }
} 