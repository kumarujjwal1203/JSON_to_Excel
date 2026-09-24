using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace GSTJsonToExcel.Models
{
    public class SelectableMonthOption : INotifyPropertyChanged
    {
        private bool _isSelected = true;
        private int _fileCount;

        public string MonthLabel { get; set; } = string.Empty;
        public int SortKey { get; set; }
        public Action<SelectableMonthOption>? OnSelectionChanged { get; set; }

        public int FileCount
        {
            get => _fileCount;
            set
            {
                if (_fileCount != value)
                {
                    _fileCount = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DisplayText));
                }
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                    OnSelectionChanged?.Invoke(this);
                }
            }
        }

        public void SetSelectedSilent(bool value)
        {
            if (_isSelected != value)
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public string DisplayText => $"{MonthLabel} ({FileCount})";

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
