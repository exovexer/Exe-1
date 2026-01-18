using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace LavoniKayitSistemi
{
    internal static class Program
    {
        private static readonly string BaseDir =
            AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        private static readonly string LogsDir = Path.Combine(BaseDir, "logs");
        private static readonly string KisilerPath = Path.Combine(LogsDir, "Kisiler.csv");
        private static readonly string SettingsPath = Path.Combine(BaseDir, "settings.json");

        private static readonly string[] KisilerFieldNames =
        {
            "Kart Numarası",
            "Ad Soyad",
            "TC Numarası",
            "Cep Telefonu Numarası",
            "İşe Başlama Tarihi",
            "Durum",
            "Adres",
            "Kan Grubu",
            "Acil Durum Kişi Adı",
            "Acil Durum Numarası",
            "Görev / Pozisyon",
            "Notlar",
            "Kayıt Tarihi",
            "Senkronize"
        };

        private const string FirebaseApiKey = "AIzaSyDBxfIpSQbZ93tIvrnf6xB2b4OQMtzrWVI";
        private const string FirebaseEmail = "uploader@eposta.com";
        private const string FirebasePassword = "A9@xF3!qW6";
        private const string FirebaseBucket = "mesaitakip-4d2d8.firebasestorage.app";
        private const int UploadDebounceSeconds = 10;

        private static readonly object DebounceLock = new();
        private static readonly Dictionary<string, Timer> DebounceTimers = new();
        private static readonly Dictionary<string, string> PendingUploadPaths = new();

        private static readonly ConcurrentDictionary<string, Dictionary<string, string>> ActiveCards = new();

        private static readonly Queue<(string Source, string CardId)> CardQueue = new();

        private static readonly HttpClient HttpClient = new();

        public static int Main(string[] args)
        {
            try
            {
                EnsureDirectoriesAndFiles();
                LoadKisiler();

                // TODO: UI, tray, and RFID keyboard listener are not yet implemented in C#.
                // The main loop below is a placeholder to keep the structure similar to the Python app.
                Console.WriteLine("Lavoni Kayıt Sistemi (C#) başlatıldı.");
                Console.WriteLine("Bu sürümde arayüz ve tray desteği henüz uygulanmadı.");

                while (true)
                {
                    ProcessCardFromQueue();
                    Thread.Sleep(200);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Uygulama hatası: {ex}");
                return 1;
            }
        }

        private static void EnsureDirectoriesAndFiles()
        {
            if (!Directory.Exists(LogsDir))
            {
                Directory.CreateDirectory(LogsDir);
            }

            if (!File.Exists(KisilerPath))
            {
                using var writer = new StreamWriter(KisilerPath, false, Encoding.UTF8);
                writer.WriteLine(string.Join(",", KisilerFieldNames));
            }
        }

        private static Dictionary<string, string> NormalizeKisiRow(Dictionary<string, string> row)
        {
            var output = new Dictionary<string, string>();
            foreach (var field in KisilerFieldNames)
            {
                if (row.TryGetValue(field, out var value) && !string.IsNullOrWhiteSpace(value))
                {
                    output[field] = value;
                }
                else
                {
                    output[field] = "0";
                }
            }

            if (!output.ContainsKey("Senkronize") || string.IsNullOrWhiteSpace(output["Senkronize"]))
            {
                output["Senkronize"] = "0";
            }

            return output;
        }

        private static void LoadKisiler()
        {
            ActiveCards.Clear();
            if (!File.Exists(KisilerPath))
            {
                return;
            }

            using var reader = new StreamReader(KisilerPath, Encoding.UTF8);
            var headerLine = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(headerLine))
            {
                return;
            }

            var headers = headerLine.Split(',');
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var parts = line.Split(',');
                var row = new Dictionary<string, string>();
                for (var i = 0; i < headers.Length && i < parts.Length; i++)
                {
                    row[headers[i]] = parts[i];
                }

                var normalized = NormalizeKisiRow(row);
                if (normalized.TryGetValue("Durum", out var durum) && durum == "Kart Aktif")
                {
                    var cardId = normalized.GetValueOrDefault("Kart Numarası", string.Empty).Trim();
                    if (!string.IsNullOrEmpty(cardId))
                    {
                        ActiveCards[cardId] = normalized;
                    }
                }
            }
        }

        private static void SaveKisilerRows(IEnumerable<Dictionary<string, string>> rows)
        {
            using var writer = new StreamWriter(KisilerPath, false, Encoding.UTF8);
            writer.WriteLine(string.Join(",", KisilerFieldNames));

            foreach (var row in rows)
            {
                var normalized = NormalizeKisiRow(row);
                var values = new List<string>();
                foreach (var field in KisilerFieldNames)
                {
                    values.Add(normalized.GetValueOrDefault(field, "0"));
                }

                writer.WriteLine(string.Join(",", values));
            }

            OnCsvWritten("KISILER", KisilerPath);
        }

        private static string GetMonthLogPath(DateTime date)
        {
            return Path.Combine(LogsDir, date.ToString("yyyy-MM", CultureInfo.InvariantCulture) + ".csv");
        }

        private static void EnsureMonthLogHeader(string path)
        {
            if (File.Exists(path))
            {
                return;
            }

            using var writer = new StreamWriter(path, false, Encoding.UTF8);
            writer.WriteLine("Tarih,Saat,Kart Numarası,Ad Soyad,Kaynak,Senkron");
        }

        private static DateTime? FindLastLogDateTime(string adSoyad, string cardId, DateTime targetMonthDate)
        {
            var logPath = GetMonthLogPath(targetMonthDate);
            if (!File.Exists(logPath))
            {
                return null;
            }

            DateTime? last = null;
            using var reader = new StreamReader(logPath, Encoding.UTF8);
            var header = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(header))
            {
                return null;
            }

            var headers = header.Split(',');
            var tarihIndex = Array.IndexOf(headers, "Tarih");
            var saatIndex = Array.IndexOf(headers, "Saat");
            var adIndex = Array.IndexOf(headers, "Ad Soyad");
            var kartIndex = Array.IndexOf(headers, "Kart Numarası");

            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var parts = line.Split(',');
                if (adIndex < 0 || kartIndex < 0 || tarihIndex < 0 || saatIndex < 0)
                {
                    continue;
                }

                if (parts.Length <= Math.Max(Math.Max(adIndex, kartIndex), Math.Max(tarihIndex, saatIndex)))
                {
                    continue;
                }

                if (!string.Equals(parts[adIndex], adSoyad, StringComparison.Ordinal) ||
                    !string.Equals(parts[kartIndex], cardId, StringComparison.Ordinal))
                {
                    continue;
                }

                if (DateTime.TryParseExact(parts[tarihIndex], "yyyy-MM-dd", CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var date) &&
                    DateTime.TryParseExact(parts[saatIndex], "HH:mm", CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var time))
                {
                    var dt = date.Date.Add(time.TimeOfDay);
                    if (last == null || dt > last)
                    {
                        last = dt;
                    }
                }
            }

            return last;
        }

        private static bool IsDuplicateLog(string adSoyad, string cardId, DateTime newDateTime, out DateTime? last)
        {
            last = FindLastLogDateTime(adSoyad, cardId, newDateTime);
            if (last == null)
            {
                return false;
            }

            var diff = newDateTime - last.Value;
            return diff >= TimeSpan.Zero && diff < TimeSpan.FromMinutes(30);
        }

        private static void AddLogAndUpdate(string adSoyad, string cardId, DateTime dt, string kaynak)
        {
            LoadKisiler();
            if (IsDuplicateLog(adSoyad, cardId, dt, out var last))
            {
                Console.WriteLine($"Duplike okutma: {adSoyad} ({cardId}) - Son okutma: {last}");
                return;
            }

            var logPath = GetMonthLogPath(dt);
            EnsureMonthLogHeader(logPath);

            using (var writer = new StreamWriter(logPath, true, Encoding.UTF8))
            {
                var row = string.Join(",",
                    dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                    dt.ToString("HH:mm", CultureInfo.InvariantCulture),
                    cardId,
                    adSoyad,
                    kaynak,
                    "0");
                writer.WriteLine(row);
            }

            OnCsvWritten("MONTH", logPath);
        }

        private static void ProcessCardFromQueue()
        {
            LoadKisiler();
            while (true)
            {
                (string Source, string CardId) entry;
                lock (CardQueue)
                {
                    if (CardQueue.Count == 0)
                    {
                        break;
                    }

                    entry = CardQueue.Dequeue();
                }

                var cardId = entry.CardId.Trim();
                if (string.IsNullOrEmpty(cardId))
                {
                    continue;
                }

                if (!ActiveCards.TryGetValue(cardId, out var person))
                {
                    continue;
                }

                var adSoyad = person.GetValueOrDefault("Ad Soyad", string.Empty).Trim();
                AddLogAndUpdate(adSoyad, cardId, DateTime.Now, entry.Source);
            }
        }

        private static void OnCsvWritten(string kind, string localPath)
        {
            lock (DebounceLock)
            {
                if (DebounceTimers.TryGetValue(kind, out var existing))
                {
                    existing.Dispose();
                }

                PendingUploadPaths[kind] = localPath;
                var timer = new Timer(_ => DebounceFire(kind), null, UploadDebounceSeconds * 1000, Timeout.Infinite);
                DebounceTimers[kind] = timer;
            }
        }

        private static void DebounceFire(string kind)
        {
            string localPath;
            lock (DebounceLock)
            {
                if (!PendingUploadPaths.TryGetValue(kind, out localPath))
                {
                    return;
                }
            }

            var remoteName = kind == "KISILER" ? "logs/Kisiler.csv" : $"logs/{Path.GetFileName(localPath)}";
            _ = Task.Run(() => UploadWithLogin(localPath, remoteName));
        }

        private static async Task UploadWithLogin(string localPath, string remoteName)
        {
            try
            {
                var token = await FirebaseLogin().ConfigureAwait(false);
                if (string.IsNullOrEmpty(token))
                {
                    return;
                }

                await FirebaseUpload(localPath, remoteName, token).ConfigureAwait(false);
            }
            catch
            {
                // Ignore upload errors to match Python behavior.
            }
        }

        private static async Task<string?> FirebaseLogin()
        {
            var payload = JsonSerializer.Serialize(new
            {
                email = FirebaseEmail,
                password = FirebasePassword,
                returnSecureToken = true
            });

            using var content = new StringContent(payload, Encoding.UTF8, "application/json");
            var url =
                $"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={FirebaseApiKey}";

            using var response = await HttpClient.PostAsync(url, content).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                return null;
            }

            var data = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            var json = JsonSerializer.Deserialize<Dictionary<string, object>>(data);
            return json != null && json.TryGetValue("idToken", out var token) ? token?.ToString() : null;
        }

        private static async Task FirebaseUpload(string localPath, string remoteName, string token)
        {
            if (!File.Exists(localPath))
            {
                return;
            }

            var body = await File.ReadAllBytesAsync(localPath).ConfigureAwait(false);
            var encodedName = Uri.EscapeDataString(remoteName);
            var url =
                $"https://firebasestorage.googleapis.com/v0/b/{FirebaseBucket}/o?uploadType=media&name={encodedName}";

            using var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new ByteArrayContent(body)
            };
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            request.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/csv");

            using var response = await HttpClient.SendAsync(request).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();
        }

        private static string FormatTurkishDateLong(DateTime date)
        {
            var months = new[]
            {
                "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
                "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
            };

            return $"{date.Day} {months[date.Month - 1]} {date.Year}";
        }

        private static DateTime ParseDottedDate(string value)
        {
            return DateTime.ParseExact(value, "dd.MM.yyyy", CultureInfo.InvariantCulture);
        }

        private static string FormatDottedDate(DateTime date)
        {
            return date.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
        }

        private static string SanitizeTurkish(string text)
        {
            var mapping = new Dictionary<char, char>
            {
                { 'ş', 's' }, { 'Ş', 'S' },
                { 'ğ', 'g' }, { 'Ğ', 'G' },
                { 'ı', 'i' }, { 'İ', 'I' },
                { 'ö', 'o' }, { 'Ö', 'O' },
                { 'ü', 'u' }, { 'Ü', 'U' },
                { 'ç', 'c' }, { 'Ç', 'C' },
            };

            var sb = new StringBuilder(text.Length);
            foreach (var ch in text)
            {
                sb.Append(mapping.TryGetValue(ch, out var replacement) ? replacement : ch);
            }

            return sb.ToString();
        }

        private sealed class AppSettings
        {
            public string ManagementPassword { get; set; } = "5353";
            public int? PopupX { get; set; }
            public int? PopupY { get; set; }
            public string PanelMode { get; set; } = "main";
        }

        private static AppSettings LoadSettings()
        {
            if (!File.Exists(SettingsPath))
            {
                return new AppSettings();
            }

            try
            {
                var json = File.ReadAllText(SettingsPath, Encoding.UTF8);
                var settings = JsonSerializer.Deserialize<AppSettings>(json);
                if (settings == null)
                {
                    return new AppSettings();
                }

                if (string.IsNullOrWhiteSpace(settings.ManagementPassword))
                {
                    settings.ManagementPassword = "5353";
                }

                if (string.IsNullOrWhiteSpace(settings.PanelMode) ||
                    (settings.PanelMode != "main" && settings.PanelMode != "management"))
                {
                    settings.PanelMode = "main";
                }

                return settings;
            }
            catch
            {
                return new AppSettings();
            }
        }

        private static void SaveSettings(AppSettings settings)
        {
            try
            {
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsPath, json, Encoding.UTF8);
            }
            catch
            {
                // Match Python behavior: ignore write errors.
            }
        }
    }
}
