using Avalonia.Collections;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Versioning;
using Windows_Startup_Manager.Models;

namespace Windows_Startup_Manager.ViewModels
{
    [SupportedOSPlatform("windows")]
    public partial class WindowsStartupManagerViewModel : ObservableObject
    {
        private readonly BackupService _backupService;

        // Collection was public; make private per MemberCanBePrivate recommendation
        private readonly ObservableCollection<StartupProcess> _items = new();
        private List<BackupService.BackupInfo> _availableBackups; // remove inline new to avoid warning

        [ObservableProperty]
        private bool _filterDisabled;

        [ObservableProperty]
        private bool _filterEnabled;

        [ObservableProperty]
        private string _searchQuery = string.Empty;

        [ObservableProperty]
        private bool _showSystemEntries = true;

        public WindowsStartupManagerViewModel()
        {
            // _startupService = new StartupService(); // removed
            _backupService = new BackupService();
            _availableBackups = new List<BackupService.BackupInfo>();

            _items.CollectionChanged += Items_CollectionChanged;

            UpdateFilteredItems();
            Refresh();
            _availableBackups = _backupService.GetAvailableBackups().ToList();
        }

        private void Items_CollectionChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            UpdateFilteredItems();
            OnPropertyChanged(nameof(FilteredItems));
        }

        public string SummaryText => $"Total {_items.Count}   Enabled {_items.Count(i => i.IsEnabled)}   Disabled {_items.Count(i => !i.IsEnabled)}";

        public ObservableCollection<StartupProcess> FilteredItems { get; } = new();

        public AvaloniaList<StartupProcess> SelectedItems { get; } = new();

        private void UpdateFilteredItems()
        {
            var filtered = _items.Where(i =>
                (string.IsNullOrWhiteSpace(SearchQuery)
                    || i.Name.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase)
                    || (!string.IsNullOrWhiteSpace(i.Publisher) && i.Publisher.Contains(SearchQuery, StringComparison.OrdinalIgnoreCase)))
                && (!FilterEnabled || i.IsEnabled)
                && (!FilterDisabled || !i.IsEnabled)
                && (ShowSystemEntries || !i.IsSystem)
            ).ToList();

            FilteredItems.Clear();
            foreach (var item in filtered)
                FilteredItems.Add(item);
        }

        [RelayCommand]
        private void Refresh()
        {
            _items.Clear();
            var processes = StartupService.GetAllStartupProcesses();
            foreach (var process in processes)
                _items.Add(process);

            UpdateFilteredItems();
            OnPropertyChanged(nameof(SummaryText));
        }

        [RelayCommand]
        private void SetSelectedStatus(bool isEnabled)
        {
            foreach (var process in SelectedItems)
            {
                if (isEnabled)
                    StartupService.EnableStartupProcess(process);
                else
                    StartupService.DisableStartupProcess(process);

                process.IsEnabled = isEnabled;
            }

            UpdateFilteredItems();
            OnPropertyChanged(nameof(SummaryText));
        }

        [RelayCommand]
        private void DeleteSelected()
        {
            foreach (var process in SelectedItems.ToList())
            {
                StartupService.RemoveStartupProcess(process);
                _items.Remove(process);
            }

            UpdateFilteredItems();
            OnPropertyChanged(nameof(SummaryText));
        }

        [RelayCommand]
        private void OpenLocations()
        {
            foreach (var process in SelectedItems)
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(process.Path))
                        Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{process.Path}\"") { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to open location for '{process.Name}': {ex.Message}");
                }
            }
        }

        [RelayCommand]
        private void Backup()
        {
            try
            {
                var path = _backupService.CreateBackup(_items);
                Debug.WriteLine($"Backup created: {path}");
                _availableBackups = _backupService.GetAvailableBackups().ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to create backup: {ex.Message}");
            }
        }

        [RelayCommand]
        private void Restore()
        {
            try
            {
                var latest = _availableBackups.OrderByDescending(b => b.Timestamp).FirstOrDefault();
                if (latest == null)
                    return;

                var data = _backupService.RestoreBackup(latest.FilePath);
                _items.Clear();
                if (data.Processes != null)
                {
                    foreach (var p in data.Processes)
                        _items.Add(p);
                }
                UpdateFilteredItems();
                OnPropertyChanged(nameof(SummaryText));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to restore backup: {ex.Message}");
            }
        }

        [RelayCommand]
        private void AddNew()
        {
            var proc = new StartupProcess("NewEntry", "Unknown", true, "Registry (Current User)", "C:/Path/To/App.exe", false);
            StartupService.AddStartupProcess(proc);
            _items.Add(proc);
            UpdateFilteredItems();
            OnPropertyChanged(nameof(SummaryText));
        }
    }
}