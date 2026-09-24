using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private readonly IFileScannerService _scannerService;
        private readonly IBatchConversionService _batchService;
        private readonly IFileService _fileService;
        private readonly ILoggingService _logger;

        private CancellationTokenSource? _conversionCts;

        // Scan state
        private BatchScanSummary? _scanSummary;
        private string _selectionTitle = "No files or folder selected";
        private string _selectionSubtitle = "Select multiple JSON files, select an entire folder, or drag & drop here.";
        private bool _isScanning;

        // Mode & Filtering
        private BatchConversionMode _conversionMode = BatchConversionMode.IndividualFiles;
        private string _searchText = string.Empty;
        private string _selectedCategoryFilter = "All";

        // Conversion progress
        private bool _isConverting;
        private int _overallProgressPercentage;
        private string _progressPhase = "Ready";
        private string _progressMessage = "Awaiting selection...";
        private string _currentFileDisplay = string.Empty;

        // Conversion results
        private bool _conversionSuccessful;
        private bool _hasError;
        private string _errorMessage = string.Empty;
        private string? _outputDirectoryPath;
        private string _elapsedTimeDisplay = string.Empty;
        private string _completionSummaryDisplay = string.Empty;

        private bool _isAnnualReportTab;

        public GSTJsonToExcel.Features.AnnualReport.ViewModels.AnnualReportViewModel AnnualReportVM { get; }

        public bool IsStandardConverterTab
        {
            get => !_isAnnualReportTab;
            set
            {
                if (value && _isAnnualReportTab)
                {
                    _isAnnualReportTab = false;
                    OnPropertyChanged(nameof(IsStandardConverterTab));
                    OnPropertyChanged(nameof(IsAnnualReportTab));
                }
            }
        }

        public bool IsAnnualReportTab
        {
            get => _isAnnualReportTab;
            set
            {
                if (SetProperty(ref _isAnnualReportTab, value))
                {
                    OnPropertyChanged(nameof(IsStandardConverterTab));
                }
            }
        }

        public ICommand SwitchToStandardTabCommand { get; }
        public ICommand SwitchToAnnualTabCommand { get; }

        public ObservableCollection<ScannedFileItem> CandidateItems { get; } = new();
        public ObservableCollection<ScannedFileItem> FilteredCandidateItems { get; } = new();
        public ObservableCollection<ScannedFileItem> IgnoredItems { get; } = new();
        public ObservableCollection<SelectableMonthOption> AvailableMonths { get; } = new();
        public ObservableCollection<GeneratedTypeOutput> GeneratedFiles { get; } = new();
        public ObservableCollection<FileProcessingErrorReport> ErrorReports { get; } = new();

        public MainViewModel(
            IFileScannerService scannerService,
            IBatchConversionService batchService,
            IFileService fileService,
            ILoggingService logger)
        {
            _scannerService = scannerService;
            _batchService = batchService;
            _fileService = fileService;
            _logger = logger;

            AnnualReportVM = new GSTJsonToExcel.Features.AnnualReport.ViewModels.AnnualReportViewModel(fileService, logger);
            SwitchToStandardTabCommand = new RelayCommand(() => IsAnnualReportTab = false);
            SwitchToAnnualTabCommand = new RelayCommand(() => IsAnnualReportTab = true);

            SelectFilesCommand = new RelayCommand(SelectFiles, () => !IsBusy);
            SelectFolderCommand = new RelayCommand(SelectFolder, () => !IsBusy);
            ConvertCommand = new RelayCommand(async () => await StartBatchConversionAsync(), () => CanConvert);
            OpenFolderCommand = new RelayCommand(OpenOutputFolder, () => CanOpenOutput);
            OpenFileCommand = new RelayCommand(param => OpenSpecificFile(param as string));
            ShowInFolderCommand = new RelayCommand(param => ShowInFolder(param as string));
            SelectAllCommand = new RelayCommand(() => SetAllSelection(true), () => !IsBusy && HasScannedFiles);
            DeselectAllCommand = new RelayCommand(() => SetAllSelection(false), () => !IsBusy && HasScannedFiles);
            SelectAllMonthsCommand = new RelayCommand(() => SetAllMonthsSelection(true), () => !IsBusy && HasAvailableMonths);
            DeselectAllMonthsCommand = new RelayCommand(() => SetAllMonthsSelection(false), () => !IsBusy && HasAvailableMonths);
            SetFilterCategoryCommand = new RelayCommand(param => SelectedCategoryFilter = param as string ?? "All");
            ClearCommand = new RelayCommand(ResetState);

            _logger.LogInfo("MainViewModel initialized with multi-file, folder scanning, and Annual Report capabilities.");
        }

        #region Properties

        public BatchScanSummary? ScanSummary
        {
            get => _scanSummary;
            set
            {
                if (SetProperty(ref _scanSummary, value))
                {
                    OnPropertyChanged(nameof(HasScannedFiles));
                    OnPropertyChanged(nameof(HasReadyToProcess));
                    OnPropertyChanged(nameof(R1Count));
                    OnPropertyChanged(nameof(R3ACount));
                    OnPropertyChanged(nameof(R2ACount));
                    OnPropertyChanged(nameof(R2BCount));
                    OnPropertyChanged(nameof(UnknownCount));
                    OnPropertyChanged(nameof(DuplicateCount));
                    OnPropertyChanged(nameof(IgnoredCount));
                    OnPropertyChanged(nameof(TotalValidCount));
                    OnPropertyChanged(nameof(TotalFilesFound));
                    OnPropertyChanged(nameof(CanConvert));
                }
            }
        }

        public string SelectionTitle
        {
            get => _selectionTitle;
            set => SetProperty(ref _selectionTitle, value);
        }

        public string SelectionSubtitle
        {
            get => _selectionSubtitle;
            set => SetProperty(ref _selectionSubtitle, value);
        }

        public bool IsScanning
        {
            get => _isScanning;
            set
            {
                if (SetProperty(ref _isScanning, value))
                {
                    OnPropertyChanged(nameof(IsBusy));
                    OnPropertyChanged(nameof(CanConvert));
                }
            }
        }

        public bool IsConverting
        {
            get => _isConverting;
            set
            {
                if (SetProperty(ref _isConverting, value))
                {
                    OnPropertyChanged(nameof(IsBusy));
                    OnPropertyChanged(nameof(CanConvert));
                }
            }
        }

        public bool IsBusy => IsScanning || IsConverting;

        public bool HasScannedFiles => ScanSummary != null && ScanSummary.TotalFilesFound > 0;
        public bool HasReadyToProcess => ScanSummary != null && ScanSummary.HasProcessableFiles;
        public bool CanConvert => HasReadyToProcess && !IsBusy;
        public bool CanOpenOutput => ConversionSuccessful && !string.IsNullOrEmpty(OutputDirectoryPath) && Directory.Exists(OutputDirectoryPath);

        // Classification Dashboard Counts
        public int R1Count => ScanSummary?.R1Count ?? 0;
        public int R3ACount => ScanSummary?.R3ACount ?? 0;
        public int R2ACount => ScanSummary?.R2ACount ?? 0;
        public int R2BCount => ScanSummary?.R2BCount ?? 0;
        public int UnknownCount => ScanSummary?.UnknownCount ?? 0;
        public int DuplicateCount => ScanSummary?.DuplicateCount ?? 0;
        public int IgnoredCount => ScanSummary?.IgnoredCount ?? 0;
        public int TotalValidCount => ScanSummary?.TotalValid ?? 0;
        public int TotalFilesFound => ScanSummary?.TotalFilesFound ?? 0;

        // Progress
        public int OverallProgressPercentage
        {
            get => _overallProgressPercentage;
            set => SetProperty(ref _overallProgressPercentage, value);
        }

        public string ProgressPhase
        {
            get => _progressPhase;
            set => SetProperty(ref _progressPhase, value);
        }

        public string ProgressMessage
        {
            get => _progressMessage;
            set => SetProperty(ref _progressMessage, value);
        }

        public string CurrentFileDisplay
        {
            get => _currentFileDisplay;
            set => SetProperty(ref _currentFileDisplay, value);
        }

        // Results
        public bool ConversionSuccessful
        {
            get => _conversionSuccessful;
            set
            {
                if (SetProperty(ref _conversionSuccessful, value))
                {
                    OnPropertyChanged(nameof(CanOpenOutput));
                }
            }
        }

        public bool HasError
        {
            get => _hasError;
            set => SetProperty(ref _hasError, value);
        }

        public string ErrorMessage
        {
            get => _errorMessage;
            set => SetProperty(ref _errorMessage, value);
        }

        public string? OutputDirectoryPath
        {
            get => _outputDirectoryPath;
            set
            {
                if (SetProperty(ref _outputDirectoryPath, value))
                {
                    OnPropertyChanged(nameof(CanOpenOutput));
                }
            }
        }

        public string ElapsedTimeDisplay
        {
            get => _elapsedTimeDisplay;
            set => SetProperty(ref _elapsedTimeDisplay, value);
        }

        public string CompletionSummaryDisplay
        {
            get => _completionSummaryDisplay;
            set => SetProperty(ref _completionSummaryDisplay, value);
        }

        // Conversion Mode
        public BatchConversionMode ConversionMode
        {
            get => _conversionMode;
            set
            {
                if (SetProperty(ref _conversionMode, value))
                {
                    OnPropertyChanged(nameof(IsIndividualMode));
                    OnPropertyChanged(nameof(IsMergedMode));
                }
            }
        }

        public bool IsIndividualMode
        {
            get => ConversionMode == BatchConversionMode.IndividualFiles;
            set
            {
                if (value) ConversionMode = BatchConversionMode.IndividualFiles;
            }
        }

        public bool IsMergedMode
        {
            get => ConversionMode == BatchConversionMode.MergedByType;
            set
            {
                if (value) ConversionMode = BatchConversionMode.MergedByType;
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value))
                {
                    ApplyFilter();
                }
            }
        }

        public string SelectedCategoryFilter
        {
            get => _selectedCategoryFilter;
            set
            {
                if (SetProperty(ref _selectedCategoryFilter, value))
                {
                    ApplyFilter();
                }
            }
        }

        public int SelectedFilesCount => CandidateItems.Count(f => f.IsReadyToProcess && f.IsSelected);
        public bool HasAvailableMonths => AvailableMonths.Count > 0;
        public string SelectedMonthsSummary => AvailableMonths.Count == 0
            ? string.Empty
            : $"{AvailableMonths.Count(m => m.IsSelected)} of {AvailableMonths.Count} Month(s) Selected  •  {SelectedFilesCount} file(s) ready to convert";

        #endregion

        #region Commands

        public ICommand SelectFilesCommand { get; }
        public ICommand SelectFolderCommand { get; }
        public ICommand ConvertCommand { get; }
        public ICommand OpenFolderCommand { get; }
        public ICommand OpenFileCommand { get; }
        public ICommand ShowInFolderCommand { get; }
        public ICommand SelectAllCommand { get; }
        public ICommand DeselectAllCommand { get; }
        public ICommand SelectAllMonthsCommand { get; }
        public ICommand DeselectAllMonthsCommand { get; }
        public ICommand SetFilterCategoryCommand { get; }
        public ICommand ClearCommand { get; }

        #endregion

        #region Methods

        public async void SelectFiles()
        {
            string[]? files = _fileService.SelectJsonFiles();
            if (files != null && files.Length > 0)
            {
                await LoadPathsAsync(files);
            }
        }

        public async void SelectFolder()
        {
            string? folder = _fileService.SelectFolder();
            if (!string.IsNullOrEmpty(folder))
            {
                await LoadPathsAsync(new[] { folder });
            }
        }

        public async Task LoadPathsAsync(IEnumerable<string> paths)
        {
            var pathList = paths.Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
            if (pathList.Count == 0) return;

            if (IsAnnualReportTab)
            {
                await AnnualReportVM.AddPathsAsync(pathList);
                return;
            }

            IsScanning = true;
            ConversionSuccessful = false;
            HasError = false;
            ErrorMessage = string.Empty;
            ProgressPhase = "Scanning Files & Folders";
            ProgressMessage = "Scanning input paths, classifying GST types, and detecting duplicates...";
            OverallProgressPercentage = 0;

            try
            {
                var summary = await _scannerService.ScanPathsAsync(pathList);
                ScanSummary = summary;

                // Update ObservableCollections for UI binding
                CandidateItems.Clear();
                foreach (var item in summary.CandidateGstFiles)
                {
                    item.PropertyChanged += (s, e) =>
                    {
                        if (e.PropertyName == nameof(ScannedFileItem.IsSelected))
                        {
                            OnPropertyChanged(nameof(SelectedFilesCount));
                            OnPropertyChanged(nameof(SelectedMonthsSummary));
                            OnPropertyChanged(nameof(CanConvert));
                        }
                    };
                    CandidateItems.Add(item);
                }

                IgnoredItems.Clear();
                foreach (var item in summary.IgnoredFiles)
                {
                    IgnoredItems.Add(item);
                }

                RebuildAvailableMonths();
                ApplyFilter();

                if (pathList.Count == 1 && Directory.Exists(pathList[0]))
                {
                    SelectionTitle = $"Folder: {Path.GetFileName(pathList[0])}";
                }
                else if (pathList.Count == 1)
                {
                    SelectionTitle = $"File: {Path.GetFileName(pathList[0])}";
                }
                else
                {
                    SelectionTitle = $"{pathList.Count} items selected";
                }

                SelectionSubtitle = $"{summary.ReadyToProcessCount} GST files ready for conversion ({summary.TotalFilesFound} total files found, {summary.IgnoredCount} non-GST ignored)";
                ProgressPhase = "Ready to Convert";
                ProgressMessage = $"{summary.ReadyToProcessCount} valid GST JSON files identified and verified.";
            }
            catch (Exception ex)
            {
                HasError = true;
                ErrorMessage = $"Scanning failed: {ex.Message}";
                _logger.LogError("File scanning error", ex);
            }
            finally
            {
                IsScanning = false;
            }
        }

        public async Task StartBatchConversionAsync()
        {
            if (ScanSummary == null || !ScanSummary.HasProcessableFiles || IsBusy) return;

            IsConverting = true;
            ConversionSuccessful = false;
            HasError = false;
            ErrorMessage = string.Empty;
            OverallProgressPercentage = 0;
            ProgressPhase = "Initializing Batch Conversion";
            ProgressMessage = ConversionMode == BatchConversionMode.IndividualFiles
                ? "Preparing individual Excel file for each JSON file..."
                : "Preparing separate output files for each GST type...";

            GeneratedFiles.Clear();
            ErrorReports.Clear();

            _conversionCts = new CancellationTokenSource();

            // Determine target output directory:
            string baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var firstFile = ScanSummary.CandidateGstFiles.FirstOrDefault(f => f.IsReadyToProcess);
            if (firstFile != null && !string.IsNullOrEmpty(firstFile.FilePath))
            {
                string? parentDir = Path.GetDirectoryName(firstFile.FilePath);
                if (!string.IsNullOrEmpty(parentDir))
                {
                    baseFolder = parentDir;
                }
            }
            string targetOutputDir = Path.Combine(baseFolder, "GST_Converted");

            var progress = new Progress<BatchConversionProgress>(p =>
            {
                OverallProgressPercentage = p.OverallPercentage;
                CurrentFileDisplay = $"{p.CurrentFileType}: {p.CurrentFileName} ({p.CurrentFileIndex}/{p.TotalFiles})";
                ProgressPhase = p.Phase;
                ProgressMessage = p.Message;
            });

            try
            {
                var result = await _batchService.ConvertBatchAsync(
                    ScanSummary,
                    targetOutputDir,
                    ConversionMode,
                    progress,
                    _conversionCts.Token);

                OutputDirectoryPath = result.OutputDirectory;
                ElapsedTimeDisplay = $"{result.ElapsedTime.TotalSeconds:F1}s";
                CompletionSummaryDisplay = result.SummaryMessage;

                foreach (var gen in result.GeneratedFiles)
                {
                    GeneratedFiles.Add(gen);
                }

                foreach (var err in result.ErrorReport)
                {
                    ErrorReports.Add(err);
                }

                ConversionSuccessful = result.Success;
                OverallProgressPercentage = 100;
                ProgressPhase = "Conversion Completed Successfully";
                ProgressMessage = $"Excel files generated in '{Path.GetFileName(targetOutputDir)}'.";
                _logger.LogInfo($"Batch conversion completed successfully: {result.SummaryMessage}");
            }
            catch (OperationCanceledException)
            {
                ProgressPhase = "Cancelled";
                ProgressMessage = "Batch conversion was cancelled by user.";
                _logger.LogInfo("Batch conversion cancelled.");
            }
            catch (Exception ex)
            {
                HasError = true;
                ErrorMessage = $"Conversion error: {ex.Message}";
                ProgressPhase = "Failed";
                ProgressMessage = ex.Message;
                _logger.LogError("Batch conversion failed", ex);
            }
            finally
            {
                IsConverting = false;
                _conversionCts?.Dispose();
                _conversionCts = null;
            }
        }

        private void RebuildAvailableMonths()
        {
            AvailableMonths.Clear();

            var readyGroups = CandidateItems
                .Where(f => f.IsReadyToProcess && !string.IsNullOrWhiteSpace(f.ReturnPeriod))
                .GroupBy(f => f.ReturnPeriod, StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g.Min(x => x.PeriodSortKey))
                .ToList();

            foreach (var group in readyGroups)
            {
                var option = new SelectableMonthOption
                {
                    MonthLabel = group.Key,
                    SortKey = group.Min(x => x.PeriodSortKey),
                    FileCount = group.Count(),
                    OnSelectionChanged = changedMonth =>
                    {
                        foreach (var item in CandidateItems.Where(f => f.IsReadyToProcess && string.Equals(f.ReturnPeriod, changedMonth.MonthLabel, StringComparison.OrdinalIgnoreCase)))
                        {
                            item.IsSelected = changedMonth.IsSelected;
                        }
                        ApplyFilter();
                        OnPropertyChanged(nameof(SelectedMonthsSummary));
                        OnPropertyChanged(nameof(CanConvert));
                    }
                };
                AvailableMonths.Add(option);
            }

            OnPropertyChanged(nameof(HasAvailableMonths));
            OnPropertyChanged(nameof(SelectedMonthsSummary));
        }

        private void SetAllMonthsSelection(bool isSelected)
        {
            foreach (var m in AvailableMonths)
            {
                m.SetSelectedSilent(isSelected);
            }
            foreach (var item in CandidateItems.Where(f => f.IsReadyToProcess))
            {
                item.IsSelected = isSelected;
            }
            ApplyFilter();
            OnPropertyChanged(nameof(SelectedMonthsSummary));
            OnPropertyChanged(nameof(CanConvert));
        }

        public void ApplyFilter()
        {
            FilteredCandidateItems.Clear();

            IEnumerable<ScannedFileItem> sourceList = SelectedCategoryFilter switch
            {
                "R1" => CandidateItems.Where(f => f.FileType == GstFileType.R1 && f.Status == FileValidationStatus.Valid),
                "R3A" => CandidateItems.Where(f => f.FileType == GstFileType.R3A && f.Status == FileValidationStatus.Valid),
                "R2A" => CandidateItems.Where(f => f.FileType == GstFileType.R2A && f.Status == FileValidationStatus.Valid),
                "R2B" => CandidateItems.Where(f => f.FileType == GstFileType.R2B && f.Status == FileValidationStatus.Valid),
                "Duplicates" => CandidateItems.Where(f => f.Status == FileValidationStatus.Duplicate),
                "Ignored" => IgnoredItems,
                _ => CandidateItems
            };

            if (AvailableMonths.Count > 0 && SelectedCategoryFilter != "Duplicates" && SelectedCategoryFilter != "Ignored")
            {
                var selectedMonthLabels = new HashSet<string>(
                    AvailableMonths.Where(m => m.IsSelected).Select(m => m.MonthLabel),
                    StringComparer.OrdinalIgnoreCase);

                sourceList = sourceList.Where(f => !f.IsReadyToProcess || selectedMonthLabels.Contains(f.ReturnPeriod));
            }

            if (!string.IsNullOrWhiteSpace(SearchText))
            {
                sourceList = sourceList.Where(f =>
                    f.FileName.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    (f.ReturnPeriod?.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ?? false));
            }

            foreach (var item in sourceList)
            {
                FilteredCandidateItems.Add(item);
            }

            OnPropertyChanged(nameof(FilteredCandidateItems));
            OnPropertyChanged(nameof(SelectedFilesCount));
            OnPropertyChanged(nameof(SelectedMonthsSummary));
        }

        private void SetAllSelection(bool isSelected)
        {
            foreach (var item in FilteredCandidateItems)
            {
                if (item.IsReadyToProcess)
                {
                    item.IsSelected = isSelected;
                }
            }
            OnPropertyChanged(nameof(SelectedFilesCount));
            OnPropertyChanged(nameof(SelectedMonthsSummary));
            OnPropertyChanged(nameof(CanConvert));
        }

        private void ShowInFolder(string? filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return;
            try
            {
                string argument = $"/select,\"{filePath}\"";
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = argument,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed revealing file in explorer: {filePath}", ex);
            }
        }

        private void OpenOutputFolder()
        {
            if (!string.IsNullOrEmpty(OutputDirectoryPath) && Directory.Exists(OutputDirectoryPath))
            {
                _fileService.OpenFolder(OutputDirectoryPath);
            }
        }

        private void OpenSpecificFile(string? filePath)
        {
            if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
            {
                _fileService.OpenFile(filePath);
            }
        }

        private void ResetState()
        {
            _conversionCts?.Cancel();
            IsScanning = false;
            IsConverting = false;

            AnnualReportVM.ClearAll();

            ScanSummary = null;
            SelectionTitle = "No files or folder selected";
            SelectionSubtitle = "Select multiple JSON files, select an entire folder, or drag & drop here.";
            CandidateItems.Clear();
            FilteredCandidateItems.Clear();
            IgnoredItems.Clear();
            AvailableMonths.Clear();
            GeneratedFiles.Clear();
            ErrorReports.Clear();
            SearchText = string.Empty;
            SelectedCategoryFilter = "All";
            ConversionSuccessful = false;
            HasError = false;
            ErrorMessage = string.Empty;
            OutputDirectoryPath = null;
            OverallProgressPercentage = 0;
            ProgressPhase = "Ready";
            ProgressMessage = "Select JSON files or a folder to begin.";
            CurrentFileDisplay = string.Empty;
            OnPropertyChanged(nameof(HasScannedFiles));
            OnPropertyChanged(nameof(HasReadyToProcess));
            OnPropertyChanged(nameof(HasAvailableMonths));
            OnPropertyChanged(nameof(SelectedMonthsSummary));
            OnPropertyChanged(nameof(SelectedFilesCount));
            OnPropertyChanged(nameof(CanConvert));
            OnPropertyChanged(nameof(CanOpenOutput));
            CommandManager.InvalidateRequerySuggested();
        }

        #endregion
    }
}
