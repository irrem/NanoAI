using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanoAI.Core
{
    public static class CommandParser
    {
        public static string Parse(string input)
        {
            input = input.ToLower();
            // Basic commands
            if (input.Contains("open notepad")) return "OpenNotepad";
            if (input.Contains("close")) return "CloseActiveWindow";
            if (input.Contains("take screenshot")) return "CaptureScreen";
            
            // FlaUI related commands
            if (input.Contains("write to notepad")) return "WriteToNotepad";
            if (input.Contains("clear text")) return "ClearNotepadText";
            if (input.Contains("append to text")) return "AppendToNotepad";
            if (input.Contains("close notepad")) return "CloseNotepad";
            
            // Parse text writing command content
            if (input.Contains("write:"))
            {
                int index = input.IndexOf("write:");
                if (index >= 0 && index + 6 < input.Length)
                {
                    string textToWrite = input.Substring(index + 6).Trim();
                    return $"WriteText:{textToWrite}";
                }
            }
            
            // If not a known command
            return "Unknown";
        }
    }
}
