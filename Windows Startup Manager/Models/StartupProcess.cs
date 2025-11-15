using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Windows_Startup_Manager.Models
{
    public class StartupProcess : INotifyPropertyChanged
    {
        private bool _isEnabled;
        private readonly bool _isSystem;
        private string _locationType;
        private string _name;
        private string _path;
        private string _publisher;

        public StartupProcess(string name, string publisher, bool isEnabled, string locationType, string path, bool isSystem)
        {
            _name = name;
            _publisher = publisher;
            _isEnabled = isEnabled;
            _locationType = locationType;
            _path = path;
            _isSystem = isSystem;
        }

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Publisher
        {
            get => _publisher;
            set
            {
                if (_publisher != value)
                {
                    _publisher = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsEnabled
        {
            get => _isEnabled;
            set
            {
                if (_isEnabled != value)
                {
                    _isEnabled = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(Status));
                }
            }
        }

        public string Status
        {
            get => IsEnabled ? "Enabled" : "Disabled";
        }

        public string LocationType
        {
            get => _locationType;
            set
            {
                if (_locationType != value)
                {
                    _locationType = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Path
        {
            get => _path;
            set
            {
                if (_path != value)
                {
                    _path = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsSystem => _isSystem;

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}