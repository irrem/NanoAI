using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.Tools;
using FlaUI.UIA3;

namespace NanAI.Automation
{
    /// <summary>
    /// Windows Hesap Makinesi uygulamasını otomatikleştirmek için sınıf
    /// </summary>
    public class CalculatorAutomation : IDisposable
    {
        private Application _app;
        private AutomationBase _automation;
        private Window _mainWindow;
        private bool _isInitialized;

        /// <summary>
        /// Hesap makinesi uygulamasını başlatır
        /// </summary>
        /// <returns>Başarılı başlatma durumu</returns>
        public async Task<bool> LaunchAsync()
        {
            try
            {
                // Hesap makinesini başlat
                _app = Application.Launch("calc.exe");
                _automation = new UIA3Automation();
                
                // Ana pencereyi bulmak için birkaç deneme
                int retryCount = 0;
                while (retryCount < 5)
                {
                    try
                    {
                        _mainWindow = _app.GetMainWindow(_automation, TimeSpan.FromSeconds(2));
                        if (_mainWindow != null)
                        {
                            _isInitialized = true;
                            return true;
                        }
                    }
                    catch
                    {
                        // Pencere henüz hazır değil, bekle ve tekrar dene
                        await Task.Delay(500);
                        retryCount++;
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hesap makinesi başlatılamadı: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Hesap makinesinde bir düğmeye basar
        /// </summary>
        /// <param name="buttonName">Düğme adı veya değeri</param>
        /// <returns>İşlem başarısı</returns>
        public bool PressButton(string buttonName)
        {
            if (!_isInitialized)
            {
                return false;
            }

            try
            {
                // Calculator açıldığında yeniden pencereyi alma
                if (_mainWindow == null || !_mainWindow.IsAvailable)
                {
                    _mainWindow = _app.GetMainWindow(_automation);
                }

                // Standart hesap makinesi tuşlarını ara
                var calcButton = FindCalculatorButton(buttonName);
                if (calcButton != null)
                {
                    calcButton.Click();
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Düğmeye basılamadı: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Hesap makinesinde bir matematik işlemi yapar
        /// </summary>
        /// <param name="expression">Matematik ifadesi (örn. "5+3")</param>
        /// <returns>İşlem sonucu</returns>
        public string Calculate(string expression)
        {
            if (!_isInitialized)
            {
                return "Hesap makinesi başlatılmadı";
            }

            try
            {
                // İfadedeki her karakteri işle
                foreach (char c in expression)
                {
                    switch (c)
                    {
                        case '+':
                            PressButton("Artı");
                            break;
                        case '-':
                            PressButton("Eksi");
                            break;
                        case '*':
                            PressButton("Çarpı");
                            break;
                        case '/':
                            PressButton("Böl");
                            break;
                        case '=':
                            PressButton("Eşittir");
                            break;
                        case '.':
                        case ',':
                            PressButton("Virgül");
                            break;
                        default:
                            if (char.IsDigit(c))
                            {
                                PressButton(c.ToString());
                            }
                            break;
                    }
                    
                    // Her tuşlama arasında kısa bekleme
                    Thread.Sleep(100);
                }
                
                // Eşittir tuşuna basılmadıysa bas
                if (!expression.Contains("="))
                {
                    PressButton("Eşittir");
                }
                
                // Sonucu al
                return GetResult();
            }
            catch (Exception ex)
            {
                return $"Hesaplama yapılamadı: {ex.Message}";
            }
        }

        /// <summary>
        /// Hesap makinesi sonucunu alır
        /// </summary>
        /// <returns>Ekrandaki sonuç</returns>
        public string GetResult()
        {
            if (!_isInitialized)
            {
                return string.Empty;
            }

            try
            {
                // Sonuç etiketini bul
                var resultText = _mainWindow.FindFirstDescendant(cf => 
                    cf.ByAutomationId("CalculatorResults"));
                
                if (resultText != null)
                {
                    // "Görüntü" kelimesini kaldır
                    string result = resultText.Name;
                    if (result.StartsWith("Görüntü"))
                    {
                        result = result.Substring(result.IndexOf(' ') + 1);
                    }
                    return result;
                }
                
                return string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Sonuç alınamadı: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Hesap makinesini kapatır
        /// </summary>
        public void Close()
        {
            if (_app != null && !_app.HasExited)
            {
                _app.Close();
                _isInitialized = false;
            }
        }

        /// <summary>
        /// Hesap makinesi düğmesini bulur
        /// </summary>
        /// <param name="buttonName">Düğme adı</param>
        /// <returns>Bulunan düğme</returns>
        private Button FindCalculatorButton(string buttonName)
        {
            // Tuş düğmeleri için sözlük
            var buttonMap = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "0", "num0Button" },
                { "1", "num1Button" },
                { "2", "num2Button" },
                { "3", "num3Button" },
                { "4", "num4Button" },
                { "5", "num5Button" },
                { "6", "num6Button" },
                { "7", "num7Button" },
                { "8", "num8Button" },
                { "9", "num9Button" },
                { "Artı", "plusButton" },
                { "Eksi", "minusButton" },
                { "Çarpı", "multiplyButton" },
                { "Böl", "divideButton" },
                { "Eşittir", "equalButton" },
                { "Virgül", "decimalSeparatorButton" },
                { "CE", "clearEntryButton" },
                { "C", "clearButton" }
            };

            // Button adından AutomationId'yi bul
            if (buttonMap.TryGetValue(buttonName, out string automationId))
            {
                return _mainWindow.FindFirstDescendant(cf => cf.ByAutomationId(automationId))?.AsButton();
            }

            // Automation ID bulunamadıysa buton adına göre ara
            return _mainWindow.FindFirstDescendant(cf => 
                cf.ByControlType(ControlType.Button).And(cf.ByName(buttonName)))?.AsButton();
        }

        /// <summary>
        /// Kaynakları temizler
        /// </summary>
        public void Dispose()
        {
            Close();
            _automation?.Dispose();
        }
    }
} 