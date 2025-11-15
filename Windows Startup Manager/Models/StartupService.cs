using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.Versioning;

namespace Windows_Startup_Manager.Models
{
    [SupportedOSPlatform("windows")]
    public class StartupService
    {
        private const string RegistryRunKeyCurrentUser = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string RegistryRunKeyLocalMachine = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string StartupFolderCurrentUser = @"Microsoft\Windows\Start Menu\Programs\Startup";
        private const string StartupFolderCommon = @"Microsoft\Windows\Start Menu\Programs\Startup";

        public static IEnumerable<StartupProcess> GetAllStartupProcesses()
        {
            var processes = new List<StartupProcess>();

            // Get processes from registry (Current User)
            processes.AddRange(GetRegistryStartupProcesses(Registry.CurrentUser, RegistryRunKeyCurrentUser, "Registry (Current User)", false));

            // Get processes from registry (Local Machine)
            processes.AddRange(GetRegistryStartupProcesses(Registry.LocalMachine, RegistryRunKeyLocalMachine, "Registry (Local Machine)", true));

            // Get processes from startup folder (Current User)
            var userStartupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "..", StartupFolderCurrentUser);
            processes.AddRange(GetFolderStartupProcesses(userStartupPath, "Startup Folder (User)", false));

            // Get processes from startup folder (Common)
            var commonStartupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "..", StartupFolderCommon);
            processes.AddRange(GetFolderStartupProcesses(commonStartupPath, "Startup Folder (Common)", true));

            return processes;
        }

        private static IEnumerable<StartupProcess> GetRegistryStartupProcesses(RegistryKey root, string keyPath, string locationType, bool isSystem)
        {
            var processes = new List<StartupProcess>();
            try
            {
                using var key = root.OpenSubKey(keyPath, false);
                if (key != null)
                {
                    foreach (var valueName in key.GetValueNames())
                    {
                        var value = key.GetValue(valueName)?.ToString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            // Extract executable path from registry value
                            var path = ExtractExecutablePath(value);
                            var publisher = GetFilePublisher(path);

                            processes.Add(new StartupProcess(
                                valueName,
                                publisher,
                                true, // Assume enabled by default
                                locationType,
                                path,
                                isSystem
                            ));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetRegistryStartupProcesses error: {ex.Message}");
            }

            return processes;
        }

        private static IEnumerable<StartupProcess> GetFolderStartupProcesses(string folderPath, string locationType, bool isSystem)
        {
            var processes = new List<StartupProcess>();
            try
            {
                if (Directory.Exists(folderPath))
                {
                    foreach (var file in Directory.GetFiles(folderPath))
                    {
                        var fileName = Path.GetFileNameWithoutExtension(file);
                        var publisher = GetFilePublisher(file);
                        processes.Add(new StartupProcess(fileName, publisher, true, locationType, file, isSystem));
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetFolderStartupProcesses error: {ex.Message}");
            }
            return processes;
        }

        public static void AddStartupProcess(StartupProcess process)
        {
            switch (process.LocationType)
            {
                case "Registry (Current User)":
                    AddToRegistry(Registry.CurrentUser, RegistryRunKeyCurrentUser, process);
                    break;
                case "Registry (Local Machine)":
                    AddToRegistry(Registry.LocalMachine, RegistryRunKeyLocalMachine, process);
                    break;
                case "Startup Folder (User)":
                    var userStartupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "..", StartupFolderCurrentUser);
                    AddToFolder(userStartupPath, process);
                    break;
                case "Startup Folder (Common)":
                    var commonStartupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "..", StartupFolderCommon);
                    AddToFolder(commonStartupPath, process);
                    break;
            }
        }

        private static void AddToRegistry(RegistryKey root, string keyPath, StartupProcess process)
        {
            try
            {
                using var key = root.CreateSubKey(keyPath, true);
                var value = process.Path;
                if (string.IsNullOrEmpty(value)) value = string.Empty;

                key.SetValue(process.Name, value);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to add startup process '{process.Name}' to registry: {ex.Message}");
            }
        }

        private static void AddToFolder(string folderPath, StartupProcess process)
        {
            try
            {
                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
                var destinationPath = Path.Combine(folderPath, Path.GetFileName(process.Path));
                if (!File.Exists(destinationPath)) File.Copy(process.Path, destinationPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"AddToFolder error: {ex.Message}");
            }
        }

        public static void RemoveStartupProcess(StartupProcess process)
        {
            switch (process.LocationType)
            {
                case "Registry (Current User)":
                    RemoveFromRegistry(Registry.CurrentUser, RegistryRunKeyCurrentUser, process);
                    break;
                case "Registry (Local Machine)":
                    RemoveFromRegistry(Registry.LocalMachine, RegistryRunKeyLocalMachine, process);
                    break;
                case "Startup Folder (User)":
                    var userStartupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "..", StartupFolderCurrentUser);
                    RemoveFromFolder(userStartupPath, process);
                    break;
                case "Startup Folder (Common)":
                    var commonStartupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "..", StartupFolderCommon);
                    RemoveFromFolder(commonStartupPath, process);
                    break;
            }
        }

        private static void RemoveFromRegistry(RegistryKey root, string keyPath, StartupProcess process)
        {
            try
            {
                using var key = root.OpenSubKey(keyPath, true);
                key?.DeleteValue(process.Name, false);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"RemoveFromRegistry error: {ex.Message}");
            }
        }

        private static void RemoveFromFolder(string folderPath, StartupProcess process)
        {
            try
            {
                var filePath = Path.Combine(folderPath, Path.GetFileName(process.Path));
                if (File.Exists(filePath)) File.Delete(filePath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"RemoveFromFolder error: {ex.Message}");
            }
        }

        public static void EnableStartupProcess(StartupProcess process) => process.IsEnabled = true;
        public static void DisableStartupProcess(StartupProcess process) => process.IsEnabled = false;

        private static string ExtractExecutablePath(string registryValue)
        {
            if (string.IsNullOrWhiteSpace(registryValue)) return registryValue;

            if (registryValue.StartsWith("\"", StringComparison.Ordinal))
            {
                var endQuote = registryValue.IndexOf("\"", 1, StringComparison.Ordinal);
                if (endQuote > 0) return registryValue[1..endQuote];
            }

            var firstSpace = registryValue.IndexOf(' ');
            if (firstSpace > 0) return registryValue[..firstSpace];

            return registryValue;
        }

        private static string GetFilePublisher(string filePath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    var versionInfo = FileVersionInfo.GetVersionInfo(filePath);
                    return string.IsNullOrWhiteSpace(versionInfo.CompanyName) ? "Unknown" : versionInfo.CompanyName!;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to read publisher from '{filePath}': {ex.Message}");
            }

            return "Unknown";
        }
    }
}