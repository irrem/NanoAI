using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using NanAI.Automation;

namespace NanAI.Gemini
{
    public class GeminiCommandInterpreter
    {
        private readonly GeminiService _geminiService;
        private readonly FormAutomation _formAutomation;
        private readonly CalculatorAutomation _calculatorAutomation;

        public GeminiCommandInterpreter(GeminiService geminiService)
        {
            _geminiService = geminiService ?? throw new ArgumentNullException(nameof(geminiService));
            _formAutomation = new FormAutomation();
            _calculatorAutomation = new CalculatorAutomation();
        }

        /// <summary>
        /// Kaynakları temizler
        /// </summary>
        public void Dispose()
        {
            _calculatorAutomation.Dispose();
        }

        /// <summary>
        /// Kullanıcı komutunu yorumlayıp uygun işlemi gerçekleştirir
        /// </summary>
        /// <param name="userCommand">Kullanıcı komutu</param>
        /// <returns>İşlem sonucu</returns>
        public async Task<string> InterpretAndExecuteCommandAsync(string userCommand)
        {
            try
            {
                // Önce Gemini'den komut yorumlamasını isteyelim
                string interpretedCommand = await _geminiService.InterpretCommandAsync(userCommand);

                // Calculator ile ilgili komutları işleyelim
                if (interpretedCommand.StartsWith("CALCULATOR:", StringComparison.OrdinalIgnoreCase))
                {
                    return HandleCalculatorCommand(interpretedCommand.Substring("CALCULATOR:".Length).Trim());
                }
                // Notepad ile ilgili komutları işleyelim
                else if (interpretedCommand.StartsWith("NOTEPAD:", StringComparison.OrdinalIgnoreCase))
                {
                    return HandleNotepadCommand(interpretedCommand.Substring("NOTEPAD:".Length).Trim());
                }
                // Uygulama başlatma komutlarını işleyelim
                else if (interpretedCommand.StartsWith("LAUNCH:", StringComparison.OrdinalIgnoreCase))
                {
                    string appName = interpretedCommand.Substring("LAUNCH:".Length).Trim();
                    return _formAutomation.LaunchApplication(appName) 
                        ? $"{appName} başlatıldı" 
                        : $"{appName} başlatılamadı";
                }
                // Gemini'ye soru sorma ve cevap alma
                else if (interpretedCommand.StartsWith("ASK:", StringComparison.OrdinalIgnoreCase))
                {
                    string question = interpretedCommand.Substring("ASK:".Length).Trim();
                    return await _geminiService.GetResponseAsync(question);
                }
                // Komut tanınmadıysa, doğrudan Gemini'den yardım alalım
                else
                {
                    return await _geminiService.GetResponseAsync(
                        $"Komut tanınmadı: {userCommand}. Hangi komutları kullanabilirim?");
                }
            }
            catch (Exception ex)
            {
                return $"Komut işleme hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Hesap makinesi komutlarını işler
        /// </summary>
        private string HandleCalculatorCommand(string command)
        {
            try
            {
                // Hesap makinesini aç
                if (command.Equals("OPEN", StringComparison.OrdinalIgnoreCase))
                {
                    return _calculatorAutomation.OpenCalculator() 
                        ? "Hesap makinesi açıldı" 
                        : "Hesap makinesi açılamadı";
                }
                // Hesap makinesini kapat
                else if (command.Equals("CLOSE", StringComparison.OrdinalIgnoreCase))
                {
                    return _calculatorAutomation.CloseCalculator() 
                        ? "Hesap makinesi kapatıldı" 
                        : "Hesap makinesi kapatılamadı";
                }
                // Matematiksel hesaplama
                else if (command.StartsWith("CALCULATE:", StringComparison.OrdinalIgnoreCase))
                {
                    string expression = command.Substring("CALCULATE:".Length).Trim();
                    // Matematiksel ifadeden gereksiz karakterleri temizle
                    expression = Regex.Replace(expression, "[^0-9+\\-*/.,()]", "");
                    return _calculatorAutomation.Calculate(expression);
                }
                else
                {
                    return $"Geçersiz hesap makinesi komutu: {command}";
                }
            }
            catch (Exception ex)
            {
                return $"Hesap makinesi komut hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Notepad komutlarını işler
        /// </summary>
        private string HandleNotepadCommand(string command)
        {
            try
            {
                // Notepad'i aç
                if (command.Equals("OPEN", StringComparison.OrdinalIgnoreCase))
                {
                    return _formAutomation.OpenNotepad() 
                        ? "Notepad açıldı" 
                        : "Notepad açılamadı";
                }
                // Notepad'e yazı yaz
                else if (command.StartsWith("WRITE:", StringComparison.OrdinalIgnoreCase))
                {
                    string text = command.Substring("WRITE:".Length).Trim();
                    return _formAutomation.WriteToNotepad(text) 
                        ? $"'{text}' Notepad'e yazıldı" 
                        : "Notepad'e yazılamadı";
                }
                else
                {
                    return $"Geçersiz Notepad komutu: {command}";
                }
            }
            catch (Exception ex)
            {
                return $"Notepad komut hatası: {ex.Message}";
            }
        }
    }
} 