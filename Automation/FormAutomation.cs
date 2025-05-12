using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace NanAI.Automation
{
    public class FormAutomation : IDisposable
    {
        private Process _process;
        private Application _application;
        private UIA3Automation _automation;
        private Window _mainWindow;
        private bool _disposed = false;

        /// <summary>
        /// Bir uygulamayı başlatır ve kontrol için hazırlar
        /// </summary>
        /// <param name="applicationPath">Uygulamanın tam yolu</param>
        /// <returns>Başarılı olup olmadığı</returns>
        public bool LaunchApplication(string applicationPath)
        {
            try
            {
                CloseApplication(); // Eğer zaten açık bir uygulama varsa kapat

                _process = Process.Start(applicationPath);
                _automation = new UIA3Automation();
                _application = new Application(_process);
                Thread.Sleep(1000); // Uygulamanın açılması için kısa bir bekleme
                _mainWindow = _application.GetMainWindow(_automation);
                
                return _mainWindow != null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Uygulama başlatma hatası: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// İsme göre uygulama başlatır
        /// </summary>
        /// <param name="applicationName">Uygulama adı (örn. "notepad", "calc")</param>
        /// <returns>Başarılı olup olmadığı</returns>
        public bool LaunchApplicationByName(string applicationName)
        {
            try
            {
                CloseApplication(); // Eğer zaten açık bir uygulama varsa kapat

                _process = Process.Start(applicationName);
                _automation = new UIA3Automation();
                _application = new Application(_process);
                Thread.Sleep(1000); // Uygulamanın açılması için kısa bir bekleme
                _mainWindow = _application.GetMainWindow(_automation);
                
                return _mainWindow != null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Uygulama başlatma hatası: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Açık uygulamayı kapatır
        /// </summary>
        public void CloseApplication()
        {
            if (_mainWindow != null && _mainWindow.IsAvailable)
            {
                _mainWindow.Close();
                _mainWindow = null;
            }

            _automation?.Dispose();
            _automation = null;

            if (_process != null && !_process.HasExited)
            {
                try
                {
                    _process.CloseMainWindow();
                    if (!_process.WaitForExit(3000))
                    {
                        _process.Kill();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Uygulama kapatma hatası: {ex.Message}");
                }
                finally
                {
                    _process = null;
                }
            }

            _application = null;
        }

        /// <summary>
        /// ID'ye göre kontrol elemanı bulur
        /// </summary>
        /// <param name="automationId">Otomasyon ID</param>
        /// <returns>Bulunan otomasyon elemanı</returns>
        public AutomationElement FindElementById(string automationId)
        {
            if (_mainWindow == null) return null;

            try
            {
                var condition = new PropertyCondition(_automation.PropertyLibrary.Element.AutomationId, automationId);
                return _mainWindow.FindFirstDescendant(condition);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Element bulma hatası: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// İsme göre kontrol elemanı bulur
        /// </summary>
        /// <param name="name">Eleman ismi</param>
        /// <returns>Bulunan otomasyon elemanı</returns>
        public AutomationElement FindElementByName(string name)
        {
            if (_mainWindow == null) return null;

            try
            {
                var condition = new PropertyCondition(_automation.PropertyLibrary.Element.Name, name);
                return _mainWindow.FindFirstDescendant(condition);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Element bulma hatası: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Kontrol türüne göre eleman bulur
        /// </summary>
        /// <param name="controlType">Kontrol türü</param>
        /// <returns>Bulunan otomasyon elemanları</returns>
        public AutomationElement[] FindElementsByControlType(ControlType controlType)
        {
            if (_mainWindow == null) return new AutomationElement[0];

            try
            {
                var condition = new PropertyCondition(_automation.PropertyLibrary.Element.ControlType, controlType);
                return _mainWindow.FindAllDescendants(condition);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Element bulma hatası: {ex.Message}");
                return new AutomationElement[0];
            }
        }

        /// <summary>
        /// Bir düğmeye tıklar
        /// </summary>
        /// <param name="buttonId">Düğme ID'si</param>
        /// <returns>Başarılı olup olmadığı</returns>
        public bool ClickButton(string buttonId)
        {
            try
            {
                var button = FindElementById(buttonId) ?? FindElementByName(buttonId);
                if (button == null) return false;

                button.Click();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Düğme tıklama hatası: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Bir metin kutusuna yazı yazar
        /// </summary>
        /// <param name="textBoxId">Metin kutusu ID'si</param>
        /// <param name="text">Yazılacak metin</param>
        /// <returns>Başarılı olup olmadığı</returns>
        public bool SetText(string textBoxId, string text)
        {
            try
            {
                var textBox = FindElementById(textBoxId) ?? FindElementByName(textBoxId);
                if (textBox == null) return false;

                textBox.Patterns.Value.Pattern.SetValue(text);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Metin yazma hatası: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Kaynak temizleme
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    CloseApplication();
                }

                _disposed = true;
            }
        }

        ~FormAutomation()
        {
            Dispose(false);
        }
    }
} 