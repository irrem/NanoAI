using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace NanoAI.Gemini
{
    /// <summary>
    /// Represents a command processed by Gemini API
    /// </summary>
    public class GeminiCommand
    {
        /// <summary>
        /// The type of command (e.g., launch, close, type, click, etc.)
        /// </summary>
        [JsonPropertyName("commandType")]
        public string CommandType { get; set; }

        /// <summary>
        /// The target of the command (application name, file path, UI element, etc.)
        /// </summary>
        [JsonPropertyName("target")]
        public string Target { get; set; }

        /// <summary>
        /// The primary action to perform (optional, for specific command types)
        /// </summary>
        [JsonPropertyName("action")]
        public string Action { get; set; }

        /// <summary>
        /// Additional parameters for the command
        /// </summary>
        [JsonPropertyName("parameters")]
        public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// Gets a parameter value as string
        /// </summary>
        /// <param name="key">Parameter key</param>
        /// <returns>Parameter value or null if not found</returns>
        public string GetParameterValue(string key)
        {
            if (Parameters != null && Parameters.TryGetValue(key, out var value))
            {
                return value?.ToString();
            }
            return null;
        }

        /// <summary>
        /// Creates a new GeminiCommand instance
        /// </summary>
        public GeminiCommand()
        {
            Parameters = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Creates a new GeminiCommand instance
        /// </summary>
        /// <param name="commandType">Command type</param>
        /// <param name="target">Command target</param>
        public GeminiCommand(string commandType, string target)
        {
            CommandType = commandType;
            Target = target;
            Parameters = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Adds a parameter to the command
        /// </summary>
        /// <param name="key">Parameter key</param>
        /// <param name="value">Parameter value</param>
        public void AddParameter(string key, string value)
        {
            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
            {
                Parameters[key] = value;
            }
        }

        /// <summary>
        /// Checks if the command is a launch command
        /// </summary>
        public bool IsLaunchCommand => CommandType?.ToLower() == "launch" || CommandType?.ToLower() == "open";

        /// <summary>
        /// Checks if the command is a close command
        /// </summary>
        public bool IsCloseCommand => CommandType?.ToLower() == "close" || CommandType?.ToLower() == "exit";

        /// <summary>
        /// Checks if the command is a type command
        /// </summary>
        public bool IsTypeCommand => CommandType?.ToLower() == "type" || CommandType?.ToLower() == "write";

        /// <summary>
        /// Checks if the command is a click command
        /// </summary>
        public bool IsClickCommand => CommandType?.ToLower() == "click" || CommandType?.ToLower() == "press";

        /// <summary>
        /// Checks if the command is an info command
        /// </summary>
        public bool IsInfoCommand => CommandType?.ToLower() == "info" || CommandType?.ToLower() == "help";

        /// <summary>
        /// Returns a string representation
        /// </summary>
        public override string ToString()
        {
            return $"{CommandType} - {Target} ({Parameters.Count} parameters)";
        }
    }
}