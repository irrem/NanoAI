using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.UIA3;
using System;
using System.Threading;

namespace NanAI.FlaUI
{
    public class NotepadAutomation
    {
        private FlaUIAutomation _automation;
        private Window _notepadWindow;
        private Application _notepadApp;

        public NotepadAutomation()
        {
            _automation = new FlaUIAutomation();
        }

        /// <summary>
        /// Launches the Notepad application
        /// </summary>
        /// <returns>Whether the operation was successful</returns>
        public bool OpenNotepad()
        {
            try
            {
                _notepadApp = Application.Launch("notepad.exe");
                Thread.Sleep(1000); // Wait for the application to start
                _notepadWindow = _notepadApp.GetMainWindow(_automation._automation);
                return _notepadWindow != null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to open Notepad: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Writes the specified text to the open Notepad
        /// </summary>
        /// <param name="text">Text to write</param>
        /// <returns>Whether the operation was successful</returns>
        public bool WriteText(string text)
        {
            try
            {
                if (_notepadWindow == null)
                {
                    Console.WriteLine("You must open Notepad first.");
                    return false;
                }

                // Find Notepad's text editor
                var document = _notepadWindow.FindFirstDescendant(cf => 
                    cf.ByControlType(ControlType.Document))?.AsTextBox();
                
                if (document == null)
                {
                    Console.WriteLine("Text document not found.");
                    return false;
                }
                
                // Write the text
                document.Text = text;
                Console.WriteLine($"Text written: {text}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write text: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Appends text to the existing content in Notepad
        /// </summary>
        /// <param name="text">Text to append</param>
        /// <returns>Whether the operation was successful</returns>
        public bool AppendText(string text)
        {
            try
            {
                if (_notepadWindow == null)
                {
                    Console.WriteLine("You must open Notepad first.");
                    return false;
                }

                // Find Notepad's text editor
                var document = _notepadWindow.FindFirstDescendant(cf =>
                    cf.ByControlType(ControlType.Document))?.AsTextBox();

                if (document == null)
                {
                    Console.WriteLine("Text document not found.");
                    return false;
                }

                // Get current text and append
                string currentText = document.Text;
                document.Text = currentText + text;
                Console.WriteLine($"Text appended: {text}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to append text: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Closes the Notepad window
        /// </summary>
        public void CloseNotepad()
        {
            try
            {
                if (_notepadWindow != null)
                {
                    _notepadApp?.Close();
                    _notepadWindow = null;
                    _notepadApp = null;
                    Console.WriteLine("Notepad closed.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while closing Notepad: {ex.Message}");
            }
        }

        /// <summary>
        /// Cleans up resources
        /// </summary>
        public void Dispose()
        {
            try
            {
                CloseNotepad();
                _automation?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while cleaning up resources: {ex.Message}");
            }
        }
    }
} 