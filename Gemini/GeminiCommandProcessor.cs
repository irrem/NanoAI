using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using NanAI.Core;
using System.Management;

namespace NanAI.Gemini
{
    /// <summary>
    /// Gemini API'den gelen komutları işleyen sınıf
    /// </summary>
    public class GeminiCommandProcessor
    {
        private readonly HttpClient _httpClient;
        private readonly string _geminiApiKey;
        private readonly string _geminiApiEndpoint;
        private readonly Dictionary<string, Func<GeminiCommand, Task<string>>> _commandHandlers;

        /// <summary>
        /// GeminiCommandProcessor sınıfının yapıcı metodu
        /// </summary>
        /// <param name="apiKey">Gemini API anahtarı</param>
        /// <param name="apiEndpoint">Gemini API uç noktası</param>
        public GeminiCommandProcessor(string apiKey, string apiEndpoint)
        {
            _httpClient = new HttpClient();
            _geminiApiKey = apiKey;
            _geminiApiEndpoint = apiEndpoint;

            // Komut işleyicilerini tanımla
            _commandHandlers = new Dictionary<string, Func<GeminiCommand, Task<string>>>(StringComparer.OrdinalIgnoreCase)
            {
                { "Calculator", ProcessCalculatorCommand },
                { "OpenApp", ProcessOpenAppCommand },
                { "ReadFile", ProcessReadFileCommand },
                { "WriteFile", ProcessWriteFileCommand },
                { "SearchWeb", ProcessSearchWebCommand },
                { "SystemInfo", ProcessSystemInfoCommand }
            };
        }

        /// <summary>
        /// Doğal dil isteğinden komut oluşturur
        /// </summary>
        /// <param name="request">Doğal dil isteği</param>
        /// <returns>Oluşturulan komut nesnesi</returns>
        public async Task<GeminiCommand> CreateCommandFromRequestAsync(string request)
        {
            try
            {
                // Gemini API'ye istek göndermek için gerekli JSON oluşturma
                var requestData = new
                {
                    contents = new[]
                    {
                        new
                        {
                            parts = new[]
                            {
                                new
                                {
                                    text = $@"
Lütfen aşağıdaki doğal dil isteğini bir komut nesnesine dönüştür:
'{request}'

Yanıtını aşağıdaki JSON formatında olmalıdır:
{{
  ""commandType"": ""Calculator"", ""OpenApp"", ""ReadFile"", ""WriteFile"", ""SearchWeb"" veya ""SystemInfo"" değerlerinden biri,
  ""target"": komutun hedefi (dosya yolu, uygulama adı, arama sorgusu vb.),
  ""parameters"": {{ isteğe bağlı parametreler ve değerleri }},
  ""description"": komutun kısa açıklaması
}}

Sadece JSON çıktısını ver, başka açıklama ekleme."
                                }
                            }
                        }
                    }
                };

                string jsonContent = JsonSerializer.Serialize(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                // API anahtarını URL'ye ekle
                string url = $"{_geminiApiEndpoint}?key={_geminiApiKey}";
                
                // API'ye istek gönder
                var response = await _httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();
                
                // Yanıtı oku
                string responseJson = await response.Content.ReadAsStringAsync();
                
                // Yanıttan JSON bölümünü çıkar
                string jsonStr = ExtractJsonFromResponse(responseJson);
                
                // JSON'ı GeminiCommand nesnesine dönüştür
                var command = JsonSerializer.Deserialize<GeminiCommand>(jsonStr) ?? new GeminiCommand();
                return command;
            }
            catch (Exception ex)
            {
                Logger.LogError($"Komut oluşturma hatası: {ex.Message}");
                return new GeminiCommand
                {
                    CommandType = "Error",
                    Description = $"İstek işlenirken hata oluştu: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Gemini API yanıtından JSON verisini çıkarır
        /// </summary>
        private string ExtractJsonFromResponse(string response)
        {
            try
            {
                // Gemini API yanıtındaki JSON yapısını ayrıştır
                var responseObject = JsonSerializer.Deserialize<JsonElement>(response);
                
                // candidates[0].content.parts[0].text içerisindeki JSON'ı bul
                string textContent = responseObject
                    .GetProperty("candidates")[0]
                    .GetProperty("content")
                    .GetProperty("parts")[0]
                    .GetProperty("text")
                    .GetString() ?? "{}";

                // Metin içinden JSON'ı çıkar (başındaki ve sonundaki metin varsa)
                int startIndex = textContent.IndexOf('{');
                int endIndex = textContent.LastIndexOf('}');
                
                if (startIndex >= 0 && endIndex > startIndex)
                {
                    return textContent.Substring(startIndex, endIndex - startIndex + 1);
                }
                
                return textContent;
            }
            catch (Exception ex)
            {
                Logger.LogError($"JSON çıkarma hatası: {ex.Message}");
                return "{}";
            }
        }

        /// <summary>
        /// Komutu işler ve sonucunu döndürür
        /// </summary>
        /// <param name="command">İşlenecek komut</param>
        /// <returns>Komut işleme sonucu</returns>
        public async Task<string> ProcessCommandAsync(GeminiCommand command)
        {
            try
            {
                // Komut türü için ilgili işleyiciyi bul
                if (_commandHandlers.TryGetValue(command.CommandType, out var handler))
                {
                    // İşleyiciyi çağır
                    return await handler(command);
                }
                
                return $"Bilinmeyen komut türü: {command.CommandType}";
            }
            catch (Exception ex)
            {
                Logger.LogError($"Komut işleme hatası: {ex.Message}");
                return $"Komut işlenirken hata oluştu: {ex.Message}";
            }
        }

        /// <summary>
        /// Hesap makinesi komutunu işler
        /// </summary>
        private async Task<string> ProcessCalculatorCommand(GeminiCommand command)
        {
            try
            {
                if (string.IsNullOrEmpty(command.Target))
                {
                    return "Hesaplanacak bir ifade belirtilmedi.";
                }

                // DataTable'ın Compute metodu ile güvenli bir şekilde ifadeyi hesapla
                var calculator = new System.Data.DataTable();
                var result = calculator.Compute(command.Target, "");
                
                return $"Sonuç: {result}";
            }
            catch (Exception ex)
            {
                return $"Hesaplama hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Uygulama açma komutunu işler
        /// </summary>
        private async Task<string> ProcessOpenAppCommand(GeminiCommand command)
        {
            try
            {
                if (string.IsNullOrEmpty(command.Target))
                {
                    return "Açılacak uygulama belirtilmedi.";
                }

                // Process sınıfını kullanarak uygulamayı başlat
                var startInfo = new ProcessStartInfo
                {
                    FileName = command.Target,
                    UseShellExecute = true
                };
                
                Process.Start(startInfo);
                return $"{command.Target} uygulaması başlatıldı.";
            }
            catch (Exception ex)
            {
                return $"Uygulama açma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Dosya okuma komutunu işler
        /// </summary>
        private async Task<string> ProcessReadFileCommand(GeminiCommand command)
        {
            try
            {
                if (string.IsNullOrEmpty(command.Target))
                {
                    return "Okunacak dosya belirtilmedi.";
                }

                // Dosyayı oku
                if (!File.Exists(command.Target))
                {
                    return $"Dosya bulunamadı: {command.Target}";
                }
                
                string content = await File.ReadAllTextAsync(command.Target);
                
                // İçerik çok uzunsa kısalt
                if (content.Length > 1000 && !command.Parameters.ContainsKey("showAll"))
                {
                    content = content.Substring(0, 1000) + "... (devamı var)";
                }
                
                return content;
            }
            catch (Exception ex)
            {
                return $"Dosya okuma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Dosyaya yazma komutunu işler
        /// </summary>
        private async Task<string> ProcessWriteFileCommand(GeminiCommand command)
        {
            try
            {
                if (string.IsNullOrEmpty(command.Target))
                {
                    return "Yazılacak dosya belirtilmedi.";
                }

                if (!command.Parameters.TryGetValue("content", out string content))
                {
                    return "Yazılacak içerik belirtilmedi.";
                }

                // Dosya içeriğine ekleme yapılacak mı?
                bool append = command.Parameters.TryGetValue("append", out string appendValue) && 
                             (appendValue.ToLower() == "true" || appendValue == "1");
                
                // Dosyaya yaz
                if (append)
                {
                    await File.AppendAllTextAsync(command.Target, content);
                    return $"{command.Target} dosyasına içerik eklendi.";
                }
                else
                {
                    await File.WriteAllTextAsync(command.Target, content);
                    return $"{command.Target} dosyasına içerik yazıldı.";
                }
            }
            catch (Exception ex)
            {
                return $"Dosya yazma hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Web arama komutunu işler
        /// </summary>
        private async Task<string> ProcessSearchWebCommand(GeminiCommand command)
        {
            try
            {
                if (string.IsNullOrEmpty(command.Target))
                {
                    return "Arama sorgusu belirtilmedi.";
                }

                // Basit bir simülasyon, gerçek bir arama motoru API'si kullanılabilir
                return $"\"{command.Target}\" için arama sonuçları gösteriliyor...";
            }
            catch (Exception ex)
            {
                return $"Web arama hatası: {ex.Message}";
            }
        }

        /// <summary>
        /// Sistem bilgisi komutunu işler
        /// </summary>
        private async Task<string> ProcessSystemInfoCommand(GeminiCommand command)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("Sistem Bilgileri:");
                sb.AppendLine($"İşletim Sistemi: {Environment.OSVersion}");
                sb.AppendLine($"Makine Adı: {Environment.MachineName}");
                sb.AppendLine($"İşlemci Sayısı: {Environment.ProcessorCount}");
                sb.AppendLine($"Sistem Dizini: {Environment.SystemDirectory}");
                sb.AppendLine($"Kullanıcı Adı: {Environment.UserName}");
                
                // WMI ile daha detaylı bilgi alma
                if (command.Parameters.TryGetValue("detail", out string detailValue) && 
                    (detailValue.ToLower() == "true" || detailValue == "1"))
                {
                    try
                    {
                        // İşlemci bilgisi
                        var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Processor");
                        foreach (var obj in searcher.Get())
                        {
                            sb.AppendLine($"CPU: {obj["Name"]}");
                            sb.AppendLine($"CPU Hızı: {obj["MaxClockSpeed"]} MHz");
                        }
                        
                        // RAM bilgisi
                        searcher = new ManagementObjectSearcher("SELECT * FROM Win32_ComputerSystem");
                        foreach (var obj in searcher.Get())
                        {
                            var totalRam = Convert.ToDouble(obj["TotalPhysicalMemory"]) / (1024 * 1024 * 1024);
                            sb.AppendLine($"Toplam RAM: {totalRam:F2} GB");
                        }
                        
                        // Disk bilgisi
                        foreach (var drive in DriveInfo.GetDrives())
                        {
                            if (drive.IsReady)
                            {
                                var totalSize = drive.TotalSize / (1024 * 1024 * 1024);
                                var freeSpace = drive.AvailableFreeSpace / (1024 * 1024 * 1024);
                                sb.AppendLine($"Disk {drive.Name} - Toplam: {totalSize:F2} GB, Boş: {freeSpace:F2} GB");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        sb.AppendLine($"Detaylı bilgi alınamadı: {ex.Message}");
                    }
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                return $"Sistem bilgisi hatası: {ex.Message}";
            }
        }
    }
} 
    }
} 