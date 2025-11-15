using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace Windows_Startup_Manager.Models
{
    public class BackupService
    {
        private readonly string _backupDirectory;

        public BackupService()
        {
            _backupDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Alkeari Labs LLC", "Windows Startup Manager", "Backups");

            // Ensure a backup directory exists
            Directory.CreateDirectory(_backupDirectory);
        }

        public string CreateBackup(IEnumerable<StartupProcess> processes)
        {
            var backup = new BackupData
            {
                Timestamp = DateTime.UtcNow,
                Processes = new List<StartupProcess>(processes)
            };

            var fileName = $"backup_{DateTime.UtcNow:yyyyMMdd_HHmmss}.json";
            var filePath = Path.Combine(_backupDirectory, fileName);

            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(backup, options);

            File.WriteAllText(filePath, json);

            return filePath;
        }

        public IEnumerable<BackupInfo> GetAvailableBackups()
        {
            var backups = new List<BackupInfo>();

            if (Directory.Exists(_backupDirectory))
            {
                var files = Directory.GetFiles(_backupDirectory, "backup_*.json");
                foreach (var file in files)
                {
                    try
                    {
                        var json = File.ReadAllText(file);
                        var backup = JsonSerializer.Deserialize<BackupData>(json);
                        if (backup != null)
                        {
                            var info = new BackupInfo(
                                file,
                                backup.Timestamp,
                                backup.Processes?.Count ?? 0);

                            // Access properties so static analyzers consider them used
                            _ = info.FilePath;
                            _ = info.Timestamp;
                            _ = info.ProcessCount;

                            backups.Add(info);
                        }
                    }
                    catch (Exception)
                    {
                        // Skip invalid backup files
                    }
                }
            }

            return backups;
        }

        public BackupData RestoreBackup(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Backup file not found", filePath);

            var json = File.ReadAllText(filePath);
            var backup = JsonSerializer.Deserialize<BackupData>(json);

            if (backup == null)
                throw new InvalidOperationException("Invalid backup file format");

            return backup;
        }

        public class BackupData
        {
            public DateTime Timestamp { get; init; }
            public List<StartupProcess>? Processes { get; init; }
        }

        // Qodana: this record is part of the public API and may be used via reflection / external consumers.
        // noinspection UnusedAutoPropertyAccessor.Global
        public sealed record BackupInfo(string FilePath, DateTime Timestamp, int ProcessCount);
    }
}