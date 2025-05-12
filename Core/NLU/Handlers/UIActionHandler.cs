using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Capturing;
using FlaUI.Core.Conditions;
using FlaUI.UIA3;
using NanoAI.Gemini;
// Alias for FlaUI ControlType to avoid ambiguity
using FlaUIControlType = FlaUI.Core.Definitions.ControlType;
// Use FlaUI AutomationElement exclusively
using AutomationElement = FlaUI.Core.AutomationElements.AutomationElement;
// Add specific references to Mouse and MouseButton
using Mouse = FlaUI.Core.Input.Mouse;
using MouseButton = FlaUI.Core.Input.MouseButton;

namespace NanoAI.Core.NLU.Handlers
{
    /// <summary>
    /// Extension methods for ConditionFactory
    /// </summary>
    public static class ConditionFactoryExtensions
    {
        /// <summary>
        /// Creates a condition that checks if a property contains a certain value
        /// </summary>
        public static ConditionBase ByNameContains(this ConditionFactory factory, string value)
        {
            // In 3.2.0, we need to use a custom comparer
            return factory.ByName(value, (FlaUI.Core.Definitions.PropertyConditionFlags)StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// Handles UI automation actions for interacting with desktop applications
    /// </summary>
    public class UIActionHandler : ICommandHandler
    {
        private Dictionary<string, Application> _runningApps = new Dictionary<string, Application>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, AutomationElement> _cachedElements = new Dictionary<string, AutomationElement>();

        public string CommandType => "ui";

        public bool CanHandle(GeminiCommand command)
        {
            return command.CommandType.Equals("ui", StringComparison.OrdinalIgnoreCase);
        }

        public async Task<CommandResult> ExecuteAsync(GeminiCommand command)
        {
            var result = new CommandResult();
            
            if (string.IsNullOrEmpty(command.Target))
            {
                result.Success = false;
                result.Message = "No target application specified for UI action";
                return result;
            }

            var action = command.Action?.ToLower();
            if (string.IsNullOrEmpty(action))
            {
                result.Success = false;
                result.Message = "No UI action specified";
                return result;
            }

            // Ensure we have a dictionary of parameters
            var parameters = command.Parameters ?? new Dictionary<string, object>();

            // If we don't have the app running, try to find and launch it
            if (!_runningApps.ContainsKey(command.Target) && action != "screenshot")
            {
                try
                {
                    var appPath = FindApplication(command.Target);
                    if (appPath == null)
                    {
                        result.Success = false;
                        result.Message = $"Could not find application '{command.Target}'";
                        return result;
                    }

                    var process = Process.Start(appPath);
                    var app = Application.Attach(process);
                    _runningApps[command.Target] = app;
                    await Task.Delay(2000); // Give the app time to initialize
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Message = $"Failed to launch or attach to application '{command.Target}': {ex.Message}";
                    return result;
                }
            }

            try
            {
                switch (action)
                {
                    case "click":
                        result = await ClickElement(command.Target, parameters);
                        break;
                    case "type":
                        result = await TypeText(command.Target, parameters);
                        break;
                    case "select":
                        result = await SelectItem(command.Target, parameters);
                        break;
                    case "rightclick":
                        result = await RightClickElement(command.Target, parameters);
                        break;
                    case "doubleclick":
                        result = await DoubleClickElement(command.Target, parameters);
                        break;
                    case "drag":
                        result = await DragElement(command.Target, parameters);
                        break;
                    case "scroll":
                        result = await ScrollElement(command.Target, parameters);
                        break;
                    case "gettext":
                        result = await GetText(command.Target, parameters);
                        break;
                    case "wait":
                        result = await WaitForElement(command.Target, parameters);
                        break;
                    case "screenshot":
                        result = await TakeScreenshot(command.Target, parameters);
                        break;
                    default:
                        result.Success = false;
                        result.Message = $"Unknown UI action: '{action}'";
                        break;
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Error performing UI action '{action}' on '{command.Target}': {ex.Message}";
            }

            return result;
        }

        private async Task<CommandResult> ClickElement(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var elementName = parameters.ContainsKey("element") ? parameters["element"]?.ToString() : null;
            
            if (string.IsNullOrEmpty(elementName))
            {
                result.Success = false;
                result.Message = "No element specified to click";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                var app = _runningApps[appName];
                var window = app.GetMainWindow(automation);
                
                var element = FindElementByName(window, elementName);
                if (element == null)
                {
                    result.Success = false;
                    result.Message = $"Could not find element '{elementName}' in application '{appName}'";
                    return result;
                }
                
                if (element.IsOffscreen)
                {
                    result.Success = false;
                    result.Message = $"Element '{elementName}' is not visible on screen";
                    return result;
                }
                
                element.Click();
                result.Message = $"Clicked on '{elementName}' in '{appName}'";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to click on element '{elementName}': {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> TypeText(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var elementName = parameters.ContainsKey("element") ? parameters["element"]?.ToString() : null;
            var text = parameters.ContainsKey("text") ? parameters["text"]?.ToString() : null;
            
            if (string.IsNullOrEmpty(text))
            {
                result.Success = false;
                result.Message = "No text specified for typing";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                
                // Check if we have a running application with this name
                if (!_runningApps.TryGetValue(appName, out var app))
                {
                    // Try to launch the application if it's not running
                    Console.WriteLine($"App {appName} not found in running apps, attempting to launch it");
                    string appPath = FindApplication(appName);
                    
                    if (!string.IsNullOrEmpty(appPath))
                    {
                        app = Application.Launch(appPath);
                        _runningApps[appName] = app;
                        // Allow time for the app to initialize
                        await Task.Delay(1000);
                    }
                    else
                    {
                        result.Success = false;
                        result.Message = $"Application '{appName}' is not running and could not be launched";
                        return result;
                    }
                }
                
                var window = app.GetMainWindow(automation);
                
                // Special handling for Notepad (which has a Document control for text)
                if (appName.Equals("notepad", StringComparison.OrdinalIgnoreCase) && string.IsNullOrEmpty(elementName))
                {
                    Console.WriteLine("Typing directly in Notepad document");
                    
                    // For Notepad, find the document control
                    var document = window.FindFirstDescendant(cf => cf.ByControlType(FlaUIControlType.Document));
                    if (document != null)
                    {
                        document.Focus();
                        document.Patterns.Value.Pattern.SetValue(text);
                        result.Message = $"Typed '{text}' into Notepad";
                        return result;
                    }
                    else
                    {
                        Console.WriteLine("Document control not found in Notepad, trying text element");
                        // If Document not found, try finding Edit control
                        var edit = window.FindFirstDescendant(cf => cf.ByControlType(FlaUIControlType.Edit));
                        if (edit != null)
                        {
                            edit.Focus();
                            edit.Patterns.Value.Pattern.SetValue(text);
                            result.Message = $"Typed '{text}' into Notepad edit control";
                            return result;
                        }
                    }
                }
                
                // Normal case - look for specific element
                AutomationElement element = null;
                
                if (!string.IsNullOrEmpty(elementName))
                {
                    // Try to find the element by name
                    element = FindElementByName(window, elementName);
                    
                    if (element == null)
                    {
                        result.Success = false;
                        result.Message = $"Could not find element '{elementName}' in application '{appName}'";
                        return result;
                    }
                }
                else
                {
                    // If no element specified, try to find a text input element that's focusable
                    var textInputs = window.FindAllDescendants(cf => 
                        cf.ByControlType(FlaUIControlType.Edit).Or(cf.ByControlType(FlaUIControlType.Document)));
                    
                    element = textInputs.FirstOrDefault(e => e.IsEnabled && !e.IsOffscreen);
                    
                    if (element == null)
                    {
                        result.Success = false;
                        result.Message = $"Could not find a suitable text input element in '{appName}'";
                        return result;
                    }
                }
                
                if (!element.Patterns.Value.IsSupported)
                {
                    result.Success = false;
                    result.Message = $"Element '{elementName}' does not support text input";
                    return result;
                }
                
                element.Focus();
                element.Patterns.Value.Pattern.SetValue(text);
                result.Message = $"Typed '{text}' into {(string.IsNullOrEmpty(elementName) ? "application" : $"'{elementName}'")} in '{appName}'";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to type into {(string.IsNullOrEmpty(elementName) ? "application" : $"element '{elementName}'")}. Error: {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> SelectItem(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var elementName = parameters.ContainsKey("element") ? parameters["element"]?.ToString() : null;
            var itemName = parameters.ContainsKey("item") ? parameters["item"]?.ToString() : null;
            
            if (string.IsNullOrEmpty(elementName))
            {
                result.Success = false;
                result.Message = "No element specified for selection";
                return result;
            }
            
            if (string.IsNullOrEmpty(itemName))
            {
                result.Success = false;
                result.Message = "No item specified for selection";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                var app = _runningApps[appName];
                var window = app.GetMainWindow(automation);
                
                var element = FindElementByName(window, elementName);
                if (element == null)
                {
                    result.Success = false;
                    result.Message = $"Could not find element '{elementName}' in application '{appName}'";
                    return result;
                }
                
                // Check if we're dealing with a list-like control
                if (TryGetAsListBox(element, out object listControl))
                {
                    // Check if we're selecting by index
                    if (int.TryParse(itemName, out int index))
                    {
                        // Select by index
                        if (listControl is ComboBox comboBox)
                        {
                            if (index < 0 || index >= comboBox.Items.Length)
                            {
                                result.Success = false;
                                result.Message = $"Index {index} is out of range (0-{comboBox.Items.Length - 1}) for combobox '{elementName}'";
                                return result;
                            }
                            comboBox.Select(index);
                        }
                        else
                        {
                            // Generic list item selection by index
                            var listItems = element.FindAllDescendants(cf => cf.ByControlType(FlaUIControlType.ListItem)).ToArray();
                            
                            if (index < 0 || index >= listItems.Length)
                            {
                                result.Success = false;
                                result.Message = $"Index {index} is out of range (0-{listItems.Length - 1}) for list '{elementName}'";
                                return result;
                            }
                            
                            if (listItems[index].Patterns.SelectionItem.IsSupported)
                            {
                                listItems[index].Patterns.SelectionItem.Pattern.Select();
                            }
                            else
                            {
                                listItems[index].Click();
                            }
                        }
                        
                        result.Message = $"Selected item at index {index} in '{elementName}' in '{appName}'";
                        return result;
                    }
                    
                    // Select by name
                    if (listControl is ComboBox combo)
                    {
                        combo.Expand();
                        await Task.Delay(500);
                        var found = false;
                        
                        foreach (var item in combo.Items)
                        {
                            if (item.Name.Equals(itemName, StringComparison.OrdinalIgnoreCase) || 
                                item.Name.IndexOf(itemName, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                item.Select();
                                found = true;
                                break;
                            }
                        }
                        
                        if (!found)
                        {
                            result.Success = false;
                            result.Message = $"Could not find item '{itemName}' in combo box '{elementName}'";
                            return result;
                        }
                    }
                    else
                    {
                        // Find the item by name inside the list
                        var listItem = element.FindFirstDescendant(cf => cf.ByName(itemName));
                        if (listItem == null)
                        {
                            // Try to find by partial name
                            var listItems = element.FindAllDescendants(cf => cf.ByControlType(FlaUIControlType.ListItem));
                            listItem = listItems.FirstOrDefault(item => item.Name.IndexOf(itemName, StringComparison.OrdinalIgnoreCase) >= 0);
                        }

                        if (listItem == null)
                        {
                            result.Success = false;
                            result.Message = $"Could not find item '{itemName}' in list '{elementName}'";
                            return result;
                        }
                        
                        if (listItem.Patterns.SelectionItem.IsSupported)
                        {
                            listItem.Patterns.SelectionItem.Pattern.Select();
                        }
                        else
                        {
                            listItem.Click();
                        }
                    }
                }
                else
                {
                    result.Success = false;
                    result.Message = $"Element '{elementName}' is not a selectable control";
                    return result;
                }
                
                result.Message = $"Selected item '{itemName}' in '{elementName}' in '{appName}'";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to select item '{itemName}' in element '{elementName}': {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> RightClickElement(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var elementName = parameters.ContainsKey("element") ? parameters["element"]?.ToString() : null;
            
            if (string.IsNullOrEmpty(elementName))
            {
                result.Success = false;
                result.Message = "No element specified to right-click";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                var app = _runningApps[appName];
                var window = app.GetMainWindow(automation);
                
                var element = FindElementByName(window, elementName);
                if (element == null)
                {
                    result.Success = false;
                    result.Message = $"Could not find element '{elementName}' in application '{appName}'";
                    return result;
                }
                
                if (element.IsOffscreen)
                {
                    result.Success = false;
                    result.Message = $"Element '{elementName}' is not visible on screen";
                    return result;
                }
                
                element.RightClick();
                result.Message = $"Right-clicked on '{elementName}' in '{appName}'";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to right-click on element '{elementName}': {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> DoubleClickElement(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var elementName = parameters.ContainsKey("element") ? parameters["element"]?.ToString() : null;
            
            if (string.IsNullOrEmpty(elementName))
            {
                result.Success = false;
                result.Message = "No element specified to double-click";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                var app = _runningApps[appName];
                var window = app.GetMainWindow(automation);
                
                var element = FindElementByName(window, elementName);
                if (element == null)
                {
                    result.Success = false;
                    result.Message = $"Could not find element '{elementName}' in application '{appName}'";
                    return result;
                }
                
                if (element.IsOffscreen)
                {
                    result.Success = false;
                    result.Message = $"Element '{elementName}' is not visible on screen";
                    return result;
                }
                
                element.DoubleClick();
                result.Message = $"Double-clicked on '{elementName}' in '{appName}'";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to double-click on element '{elementName}': {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> DragElement(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var sourceElementName = parameters.ContainsKey("source") ? parameters["source"]?.ToString() : null;
            var targetElementName = parameters.ContainsKey("target") ? parameters["target"]?.ToString() : null;
            
            if (string.IsNullOrEmpty(sourceElementName))
            {
                result.Success = false;
                result.Message = "No source element specified for drag operation";
                return result;
            }
            
            if (string.IsNullOrEmpty(targetElementName))
            {
                result.Success = false;
                result.Message = "No target element specified for drag operation";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                var app = _runningApps[appName];
                var window = app.GetMainWindow(automation);
                
                var sourceElement = FindElementByName(window, sourceElementName);
                if (sourceElement == null)
                {
                    result.Success = false;
                    result.Message = $"Could not find source element '{sourceElementName}' in application '{appName}'";
                    return result;
                }
                
                var targetElement = FindElementByName(window, targetElementName);
                if (targetElement == null)
                {
                    result.Success = false;
                    result.Message = $"Could not find target element '{targetElementName}' in application '{appName}'";
                    return result;
                }
                
                if (sourceElement.IsOffscreen || targetElement.IsOffscreen)
                {
                    result.Success = false;
                    result.Message = $"Either source or target element is not visible on screen";
                    return result;
                }
                
                var sourceCenter = sourceElement.GetClickablePoint();
                var targetCenter = targetElement.GetClickablePoint();
                
                Mouse.Position = sourceCenter;
                Mouse.Down(MouseButton.Left);
                await Task.Delay(500);
                Mouse.Position = targetCenter;
                await Task.Delay(500);
                Mouse.Up(MouseButton.Left);
                
                result.Message = $"Dragged '{sourceElementName}' to '{targetElementName}' in '{appName}'";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to drag from '{sourceElementName}' to '{targetElementName}': {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> ScrollElement(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var elementName = parameters.ContainsKey("element") ? parameters["element"]?.ToString() : null;
            var direction = parameters.ContainsKey("direction") ? parameters["direction"]?.ToString()?.ToLower() : "down";
            
            // Parse the amount parameter safely
            int amount = 3; // Default value
            if (parameters.ContainsKey("amount"))
            {
                var amountStr = parameters["amount"]?.ToString();
                if (!int.TryParse(amountStr, out amount))
                {
                    amount = 3; // Default if parsing fails
                }
            }
            
            if (string.IsNullOrEmpty(elementName))
            {
                result.Success = false;
                result.Message = "No element specified to scroll";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                var app = _runningApps[appName];
                var window = app.GetMainWindow(automation);
                
                var element = FindElementByName(window, elementName);
                if (element == null)
                {
                    result.Success = false;
                    result.Message = $"Could not find element '{elementName}' in application '{appName}'";
                    return result;
                }
                
                if (element.IsOffscreen)
                {
                    result.Success = false;
                    result.Message = $"Element '{elementName}' is not visible on screen";
                    return result;
                }
                
                var clickPoint = element.GetClickablePoint();
                Mouse.Position = clickPoint;
                
                for (int i = 0; i < amount; i++)
                {
                    if (direction == "up")
                        Mouse.Scroll(1);
                    else if (direction == "down")
                        Mouse.Scroll(-1);
                    else if (direction == "left")
                        Mouse.HorizontalScroll(1);
                    else if (direction == "right")
                        Mouse.HorizontalScroll(-1);
                    
                    await Task.Delay(100);
                }
                
                result.Message = $"Scrolled {direction} in '{elementName}' in '{appName}'";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to scroll {direction} in element '{elementName}': {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> GetText(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var elementName = parameters.ContainsKey("element") ? parameters["element"]?.ToString() : null;
            
            if (string.IsNullOrEmpty(elementName))
            {
                result.Success = false;
                result.Message = "No element specified to get text from";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                var app = _runningApps[appName];
                var window = app.GetMainWindow(automation);
                
                var element = FindElementByName(window, elementName);
                if (element == null)
                {
                    result.Success = false;
                    result.Message = $"Could not find element '{elementName}' in application '{appName}'";
                    return result;
                }
                
                string text = null;
                
                if (element.Patterns.Value.IsSupported)
                {
                    text = element.Patterns.Value.Pattern.Value;
                }
                else
                {
                    text = element.Name;
                }
                
                result.Message = $"Text from '{elementName}' in '{appName}': {text}";
                result.AdditionalData["text"] = text;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to get text from element '{elementName}': {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> WaitForElement(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var elementName = parameters.ContainsKey("element") ? parameters["element"]?.ToString() : null;
            
            // Parse the timeout parameter safely
            int timeout = 10; // Default value
            if (parameters.ContainsKey("timeout"))
            {
                var timeoutStr = parameters["timeout"]?.ToString();
                if (!int.TryParse(timeoutStr, out timeout))
                {
                    timeout = 10; // Default if parsing fails
                }
            }
            
            if (string.IsNullOrEmpty(elementName))
            {
                result.Success = false;
                result.Message = "No element specified to wait for";
                return result;
            }
            
            try
            {
                var automation = new UIA3Automation();
                var app = _runningApps[appName];
                var window = app.GetMainWindow(automation);
                
                var stopwatch = Stopwatch.StartNew();
                AutomationElement element = null;
                
                while (stopwatch.Elapsed.TotalSeconds < timeout)
                {
                    element = FindElementByName(window, elementName);
                    if (element != null && !element.IsOffscreen)
                        break;
                    
                    await Task.Delay(500);
                }
                
                stopwatch.Stop();
                
                if (element == null || element.IsOffscreen)
                {
                    result.Success = false;
                    result.Message = $"Element '{elementName}' did not appear within {timeout} seconds";
                    return result;
                }
                
                result.Message = $"Element '{elementName}' appeared after {stopwatch.Elapsed.TotalSeconds:0.0} seconds in '{appName}'";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed while waiting for element '{elementName}': {ex.Message}";
            }
            
            return result;
        }
        
        private async Task<CommandResult> TakeScreenshot(string appName, Dictionary<string, object> parameters)
        {
            var result = new CommandResult();
            var savePath = parameters.ContainsKey("path") ? parameters["path"]?.ToString() : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "screenshot.png");
            
            try
            {
                if (string.IsNullOrEmpty(appName) || appName.Equals("desktop", StringComparison.OrdinalIgnoreCase))
                {
                    // Take screenshot of entire desktop
                    var screenshot = Capture.Screen();
                    var dir = Path.GetDirectoryName(savePath);
                    if (!Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                    
                    screenshot.Bitmap.Save(savePath);
                    result.Message = $"Captured screenshot of desktop and saved to {savePath}";
                    result.AdditionalData["screenshotPath"] = savePath;
                }
                else
                {
                    // Take screenshot of specific app
                    if (!_runningApps.ContainsKey(appName))
                    {
                        result.Success = false;
                        result.Message = $"Application '{appName}' is not running";
                        return result;
                    }
                    
                    var automation = new UIA3Automation();
                    var app = _runningApps[appName];
                    var window = app.GetMainWindow(automation);
                    
                    var rectangle = window.BoundingRectangle;
                    var screenshot = Capture.Rectangle(new Rectangle((int)rectangle.X, (int)rectangle.Y, (int)rectangle.Width, (int)rectangle.Height));
                    
                    var dir = Path.GetDirectoryName(savePath);
                    if (!Directory.Exists(dir))
                        Directory.CreateDirectory(dir);
                    
                    screenshot.Bitmap.Save(savePath);
                    result.Message = $"Captured screenshot of '{appName}' and saved to {savePath}";
                    result.AdditionalData["screenshotPath"] = savePath;
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to capture screenshot: {ex.Message}";
            }
            
            return result;
        }
        
        private AutomationElement FindElementByName(FlaUI.Core.AutomationElements.Window window, string name)
        {
            // First try exact match
            var element = window.FindFirstDescendant(cf => cf.ByName(name));
            
            // If not found, try contains match
            if (element == null)
            {
                // Try to find element with partial name match
                var allElements = window.FindAllDescendants();
                element = allElements.FirstOrDefault(e => 
                    !string.IsNullOrEmpty(e.Name) && 
                    e.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            
            // If still not found, try by AutomationId
            if (element == null)
            {
                element = window.FindFirstDescendant(cf => cf.ByAutomationId(name));
            }
            
            return element;
        }
        
        /// <summary>
        /// Try to convert an element to a proper ListBox
        /// </summary>
        private bool TryGetAsListBox(AutomationElement element, out object listControl)
        {
            listControl = null;
            
            try
            {
                if (element.ControlType == FlaUIControlType.List)
                {
                    // Default fallback
                    listControl = element;
                    return true;
                }
                else if (element.ControlType == FlaUIControlType.ComboBox)
                {
                    listControl = element.AsComboBox();
                    return true;
                }
                else if (element.ControlType == FlaUIControlType.Tree)
                {
                    listControl = element.AsTree();
                    return true; 
                }
                else if (element.ControlType == FlaUIControlType.DataGrid)
                {
                    listControl = element.AsGrid();
                    return true;
                }
                else if (element.ControlType == FlaUIControlType.ListItem)
                {
                    // Parent element should be a list
                    var parent = element.Parent;
                    if (parent != null)
                    {
                        return TryGetAsListBox(parent, out listControl);
                    }
                }
            }
            catch (Exception)
            {
                // If conversion fails, return false
                return false;
            }
            
            return false;
        }
        
        private string FindApplication(string appName)
        {
            // Common paths to check
            var commonPaths = new List<string>
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                @"C:\Windows",
                @"C:\Windows\System32"
            };
            
            // Check if the app name has an extension
            if (!appName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                appName = appName + ".exe";
            
            // First check if the app is already running
            var processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(appName));
            if (processes.Length > 0)
            {
                return processes[0].MainModule.FileName;
            }
            
            // Then check common paths
            foreach (var basePath in commonPaths)
            {
                var files = Directory.GetFiles(basePath, appName, SearchOption.AllDirectories);
                if (files.Length > 0)
                    return files[0];
                
                // Also check in subdirectories with the app name
                var appNameWithoutExt = Path.GetFileNameWithoutExtension(appName);
                var possibleDirs = Directory.GetDirectories(basePath, appNameWithoutExt, SearchOption.AllDirectories);
                foreach (var dir in possibleDirs)
                {
                    files = Directory.GetFiles(dir, appName, SearchOption.AllDirectories);
                    if (files.Length > 0)
                        return files[0];
                }
            }
            
            return null;
        }
    }
} 
