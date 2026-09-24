using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;

namespace GSTJsonToExcel.Models
{
    public class ScannedFileItem : INotifyPropertyChanged
    {
        private bool _isSelected = true;
        private ProcessingState _processingState = ProcessingState.Pending;
        private string? _statusMessage;
        private int _recordsExported;
        private string? _outputExcelPath;

        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public long FileSizeBytes { get; set; }
        public string Extension { get; set; } = string.Empty;
        public GstFileType FileType { get; set; } = GstFileType.Unknown;
        public string ReturnPeriod { get; set; } = "Unknown";
        public int PeriodSortKey { get; set; } = 999999;
        public FileValidationStatus Status { get; set; } = FileValidationStatus.Valid;
        public string? ContentHash { get; set; }
        public string? DuplicateOf { get; set; }
        public string? StatusReason { get; set; }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        public ProcessingState ProcessingState
        {
            get => _processingState;
            set
            {
                if (_processingState != value)
                {
                    _processingState = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(ProcessingStateDisplay));
                }
            }
        }

        public string ProcessingStateDisplay => ProcessingState switch
        {
            ProcessingState.Pending => "Pending",
            ProcessingState.Converting => "Converting...",
            ProcessingState.Completed => "Completed",
            ProcessingState.Failed => "Failed",
            ProcessingState.Skipped => "Skipped",
            _ => "Unknown"
        };

        public string? StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        public int RecordsExported
        {
            get => _recordsExported;
            set
            {
                if (_recordsExported != value)
                {
                    _recordsExported = value;
                    OnPropertyChanged();
                }
            }
        }

        public string? OutputExcelPath
        {
            get => _outputExcelPath;
            set
            {
                if (_outputExcelPath != value)
                {
                    _outputExcelPath = value;
                    OnPropertyChanged();
                }
            }
        }

        public string TargetExcelFileName
        {
            get
            {
                if (string.IsNullOrEmpty(FileName)) return "output.xlsx";
                return Path.GetFileNameWithoutExtension(FileName) + ".xlsx";
            }
        }

        public string FileSizeFormatted
        {
            get
            {
                double kb = (double)FileSizeBytes / 1024.0;
                return kb > 1024 ? $"{kb / 1024.0:F2} MB" : $"{kb:F1} KB";
            }
        }

        public bool IsReadyToProcess => Status == FileValidationStatus.Valid && FileType != GstFileType.Unknown;

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
