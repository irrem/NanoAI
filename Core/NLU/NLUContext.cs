using System;
using System.Collections.Generic;
using NanoAI.Gemini;

namespace NanoAI.Core.NLU
{
    /// <summary>
    /// Manages conversation state and context for NLU processing
    /// </summary>
    public class NLUContext
    {
        /// <summary>
        /// Session identifier
        /// </summary>
        public string SessionId { get; set; } = Guid.NewGuid().ToString();
        
        /// <summary>
        /// User identifier
        /// </summary>
        public string UserId { get; set; } = "default_user";
        
        /// <summary>
        /// Conversation history
        /// </summary>
        public List<string> ConversationHistory { get; set; } = new List<string>();
        
        /// <summary>
        /// Variables stored in the context
        /// </summary>
        public Dictionary<string, object> Variables { get; set; } = new Dictionary<string, object>();
        
        /// <summary>
        /// Indicates if the current command is part of a composite command
        /// </summary>
        public bool IsPartOfCompositeCommand { get; set; } = false;
        
        /// <summary>
        /// The parent command if this is part of a composite command
        /// </summary>
        public string ParentCommand { get; set; }

        public List<CommandHistory> CommandHistory { get; set; }
        public Dictionary<string, object> SessionVariables { get; set; }
        public DateTime LastInteraction { get; set; }

        public NLUContext()
        {
            CommandHistory = new List<CommandHistory>();
            SessionVariables = new Dictionary<string, object>();
            LastInteraction = DateTime.UtcNow;
        }

        public void AddCommandToHistory(string originalCommand, GeminiCommand processedCommand, string result)
        {
            CommandHistory.Add(new CommandHistory
            {
                OriginalCommand = originalCommand,
                ProcessedCommand = processedCommand,
                Result = result,
                Timestamp = DateTime.UtcNow
            });
            LastInteraction = DateTime.UtcNow;
        }

        public void SetVariable(string key, object value)
        {
            SessionVariables[key] = value;
        }

        public T GetVariable<T>(string key, T defaultValue = default)
        {
            if (SessionVariables.TryGetValue(key, out var value) && value is T typedValue)
            {
                return typedValue;
            }
            return defaultValue;
        }

        public void ClearHistory()
        {
            CommandHistory.Clear();
        }

        public void ClearVariables()
        {
            SessionVariables.Clear();
        }
    }

    public class CommandHistory
    {
        public string OriginalCommand { get; set; }
        public GeminiCommand ProcessedCommand { get; set; }
        public string Result { get; set; }
        public DateTime Timestamp { get; set; }
    }
} 