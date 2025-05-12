using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using NanoAI.Gemini;

namespace NanoAI.Core.NLU
{
    /// <summary>
    /// Interface for command handlers that process specific command types
    /// </summary>
    public interface ICommandHandler
    {
        bool CanHandle(GeminiCommand command);
        Task<CommandResult> ExecuteAsync(GeminiCommand command);
    }

    /// <summary>
    /// Result of command execution
    /// </summary>
    public class CommandResult
    {
        public string Message { get; set; }
        public bool Success { get; set; } = true;
        public System.Collections.Generic.List<string> Suggestions { get; set; } = new System.Collections.Generic.List<string>();
        public Dictionary<string, object> AdditionalData { get; set; }

        public CommandResult()
        {
            AdditionalData = new Dictionary<string, object>();
        }

        public CommandResult(bool success, string message)
        {
            Success = success;
            Message = message;
            AdditionalData = new Dictionary<string, object>();
        }

        public static CommandResult SuccessResult(string message = null)
        {
            return new CommandResult
            {
                Success = true,
                Message = message
            };
        }

        public static CommandResult ErrorResult(string message, List<string> suggestions = null)
        {
            return new CommandResult
            {
                Success = false,
                Message = message,
                Suggestions = suggestions ?? new List<string>()
            };
        }
    }
} 