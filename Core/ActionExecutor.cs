using NanAI.Input;
using NanAI.FlaUI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanAI.Core
{
    public static class ActionExecutor
    {
        private static NotepadAutomation _notepadAutomation;

        static ActionExecutor()
        {
            _notepadAutomation = new NotepadAutomation();
        }

        public static void Execute(string command)
        {
            if (command.StartsWith("WriteText:"))
            {
                string textToWrite = command.Substring("WriteText:".Length);
                WriteTextToNotepad(textToWrite);
                return;
            }

            switch (command)
            {
                case "OpenNotepad":
                    Process.Start("notepad.exe");
                    break;
                case "CloseActiveWindow":
                    InputController.CloseActiveWindow();
                    break;
                case "CaptureScreen":
                    VisionService.CaptureAndSaveScreen();
                    break;
                // FlaUI related commands
                case "WriteToNotepad":
                    OpenNotepadWithFlaUI();                  
                    break;
                case "ClearNotepadText":
                    _notepadAutomation.WriteText("");
                    break;
                case "AppendToNotepad":
                    _notepadAutomation.AppendText("\r\nThis text was appended to the existing text.");
                    break;
                case "CloseNotepad":
                    _notepadAutomation.CloseNotepad();
                    break;
                default:
                    Console.WriteLine("Command not understood.");
                    break;
            }
        }

        private static void OpenNotepadWithFlaUI()
        {
            // Don't open Notepad if it's already open
            bool result = _notepadAutomation.OpenNotepad();
            if (result)
            {
                Console.WriteLine("Notepad opened with FlaUI.");
            }
            else
            {
                Console.WriteLine("Failed to open Notepad, retrying...");
                _notepadAutomation = new NotepadAutomation();
                _notepadAutomation.OpenNotepad();
            }
        }

        private static void WriteTextToNotepad(string text)
        {
            OpenNotepadWithFlaUI();
            bool result = _notepadAutomation.WriteText(text);
            if (result)
            {
                Console.WriteLine($"Text written: {text}");
            }
            else
            {
                Console.WriteLine("Failed to write text.");
            }
        }

        public static void CleanUp()
        {
            _notepadAutomation?.Dispose();
        }
    }
}
