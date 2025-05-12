using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using System;
using System.Threading;

namespace NanAI.FlaUI
{
    public class FlaUIAutomation
    {
        public UIA3Automation _automation;

        public FlaUIAutomation()
        {
            _automation = new UIA3Automation();
        }

        /// <summary>
        /// Launches the given application and returns its main window
        /// </summary>
        /// <param name="appName">Name of the application to launch (e.g., "notepad.exe")</param>
        /// <returns>Main window of the application</returns>
        public Window LaunchApplication(string appName)
        {
            try
            {
                var app = Application.Launch(appName);
                Thread.Sleep(1000); // Wait for the application to start
                var window = app.GetMainWindow(_automation);
                Console.WriteLine($"Application launched: {window.Title}");
                return window;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to launch application: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Connects to an already running application
        /// </summary>
        /// <param name="processName">Process name (e.g., "notepad")</param>
        /// <returns>Main window of the application</returns>
        public Window AttachToApplication(string processName)
        {
            try
            {
                var app = Application.Attach(processName);
                var window = app.GetMainWindow(_automation);
                Console.WriteLine($"Connected to application: {window.Title}");
                return window;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to connect to application: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Cleans up resources
        /// </summary>
        public void Dispose()
        {
            _automation?.Dispose();
        }
    }
} 