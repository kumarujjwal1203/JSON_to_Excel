using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Input;
using GSTJsonToExcel.Features.AnnualReport.Models;
using GSTJsonToExcel.Features.AnnualReport.Services;
using GSTJsonToExcel.Services.Interfaces;
using GSTJsonToExcel.ViewModels;

namespace GSTJsonToExcel.Features.AnnualReport.ViewModels
{
    public class AnnualReportViewModel : ViewModelBase
    {
        public const string AllUploadedPeriodsOption = "All Uploaded Months (Auto)";

        private readonly IFileService _fileService;
        private readonly ILoggingService _logger;
        private readonly AnnualReportScannerService _scanner = new();
        private readonly AnnualReportAggregatorService _aggregator = new();
        private readonly AnnualReportExcelBuilder _excelBuilder = new();

        private CancellationTokenSource? _cts;
        private List<AnnualScannedFile> _allScannedFiles = new();

        private string _companyName = string.Empty;
        private string _selectedFinancialYear = AllUploadedPeriodsOption;
        private bool _forceFull12MonthColumns;
        private GstinDisplayMode _gstinDisplayMode = GstinDisplayMode.ShowGstinAndState;
        private bool _separateWorkbookPerGstin;
        private string _outputFolderPath = string.Empty;

        private bool _includeOverview = true;
        private bool _includeGstr3B = true;
        private bool _includeGstr1 = true;
        private bool _includeGstr2A = true;
        private bool _includeGstr2B = true;
        private bool _include3BVs1 = true;
        private bool _include3BVs2A = true;
        private bool _include3BVs2B = true;

        private bool _canInclude3B;
        private bool _canInclude1;
        private bool _canInclude2A;
        private bool _canInclude2B;
        private bool _canInclude3BVs1;
        private bool _canInclude3BVs2A;
        private bool _canInclude3BVs2B;

        private bool _isBusy;
        private int _progressPercentage;
        private string _progressStatusText = "Ready to scan GST files";
        private string _scanBadgeText = "Drop or select 4 files (1 month's GSTR-1/3B/2A/2B), multiple months, or full annual files";
        private bool _hasGenerationResult;
        private string _includedSheetsSummary = string.Empty;
        private string _skippedSheetsSummary = string.Empty;
        private string _errorBannerText = string.Empty;

        public ObservableCollection<string> AvailableFinancialYears { get; } = new();
        public ObservableCollection<GSTJsonToExcel.Models.SelectableMonthOption> AvailableMonths { get; } = new();
        public ObservableCollection<GstinMonthCoverageRow> CoverageRows { get; } = new();
        public ObservableCollection<GstinCustomLabelItem> CustomLabelItems { get; } = new();
        public ObservableCollection<string> GeneratedFiles { get; } = new();
        public ObservableCollection<AnnualReconciliationAlert> ReconciliationAlerts { get; } = new();

        public AnnualReportViewModel(IFileService fileService, ILoggingService logger)
        {
            _fileService = fileService;
            _logger = logger;

            LoadPersistedSettings();

            AddFilesCommand = new RelayCommand(async () => await SelectFilesAsync(), () => !IsBusy);
            AddFolderCommand = new RelayCommand(async () => await SelectFolderAsync(), () => !IsBusy);
            ClearFilesCommand = new RelayCommand(ClearAll, () => !IsBusy);
            BrowseOutputFolderCommand = new RelayCommand(BrowseOutputFolder, () => !IsBusy);
            GenerateAnnualReportCommand = new RelayCommand(async () => await GenerateReportAsync(), () => CanGenerate);
            CancelCommand = new RelayCommand(() => _cts?.Cancel(), () => IsBusy);
            OpenFileCommand = new RelayCommand(p => OpenFile(p as string));
            ShowInFolderCommand = new RelayCommand(p => ShowInFolder(p as string));
            SelectAllMonthsCommand = new RelayCommand(() => SetAllMonthsSelection(true), () => !IsBusy && HasAvailableMonths);
            DeselectAllMonthsCommand = new RelayCommand(() => SetAllMonthsSelection(false), () => !IsBusy && HasAvailableMonths);
        }

        #region Commands

        public ICommand AddFilesCommand { get; }
        public ICommand AddFolderCommand { get; }
        public ICommand ClearFilesCommand { get; }
        public ICommand BrowseOutputFolderCommand { get; }
        public ICommand GenerateAnnualReportCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand OpenFileCommand { get; }
        public ICommand ShowInFolderCommand { get; }
        public ICommand SelectAllMonthsCommand { get; }
        public ICommand DeselectAllMonthsCommand { get; }

        #endregion

        #region Bindable Properties

        public bool HasScannedFiles => _allScannedFiles.Count > 0;
        public bool HasAvailableMonths => AvailableMonths.Count > 0;
        public string SelectedMonthsSummary => AvailableMonths.Count == 0
            ? "Upload JSON files to choose specific months"
            : $"{AvailableMonths.Count(m => m.IsSelected)} of {AvailableMonths.Count} Month(s) Selected ({GetFilteredValidFiles().Count} files)";
        public bool CanGenerate => !IsBusy && HasScannedFiles && !string.IsNullOrWhiteSpace(SelectedFinancialYear) && AvailableMonths.Any(m => m.IsSelected);

        public string CompanyName
        {
            get => _companyName;
            set
            {
                if (SetProperty(ref _companyName, value))
                {
                    SavePersistedSettings();
                }
            }
        }

        public string SelectedFinancialYear
        {
            get => _selectedFinancialYear;
            set
            {
                if (SetProperty(ref _selectedFinancialYear, value))
                {
                    RebuildAvailableMonthsForSelectedFy();
                    RefreshCoverageForSelectedFy();
                    OnPropertyChanged(nameof(CanGenerate));
                }
            }
        }

        public bool ForceFull12MonthColumns
        {
            get => _forceFull12MonthColumns;
            set
            {
                if (SetProperty(ref _forceFull12MonthColumns, value))
                {
                    RefreshCoverageForSelectedFy();
                    SavePersistedSettings();
                }
            }
        }

        public bool IsShowGstinAndState
        {
            get => _gstinDisplayMode == GstinDisplayMode.ShowGstinAndState;
            set
            {
                if (value)
                {
                    _gstinDisplayMode = GstinDisplayMode.ShowGstinAndState;
                    NotifyDisplayModeChanged();
                }
            }
        }

        public bool IsShowStateOnly
        {
            get => _gstinDisplayMode == GstinDisplayMode.ShowStateOnly;
            set
            {
                if (value)
                {
                    _gstinDisplayMode = GstinDisplayMode.ShowStateOnly;
                    NotifyDisplayModeChanged();
                }
            }
        }

        public bool IsShowCustomBranchName
        {
            get => _gstinDisplayMode == GstinDisplayMode.ShowCustomBranchName;
            set
            {
                if (value)
                {
                    _gstinDisplayMode = GstinDisplayMode.ShowCustomBranchName;
                    NotifyDisplayModeChanged();
                }
            }
        }

        private void NotifyDisplayModeChanged()
        {
            OnPropertyChanged(nameof(IsShowGstinAndState));
            OnPropertyChanged(nameof(IsShowStateOnly));
            OnPropertyChanged(nameof(IsShowCustomBranchName));
            SavePersistedSettings();
        }

        public bool SeparateWorkbookPerGstin
        {
            get => _separateWorkbookPerGstin;
            set
            {
                if (SetProperty(ref _separateWorkbookPerGstin, value))
                {
                    SavePersistedSettings();
                }
            }
        }

        public string OutputFolderPath
        {
            get => _outputFolderPath;
            set
            {
                if (SetProperty(ref _outputFolderPath, value))
                {
                    SavePersistedSettings();
                }
            }
        }

        public bool IncludeOverview
        {
            get => _includeOverview;
            set => SetProperty(ref _includeOverview, value);
        }

        public bool IncludeGstr3B
        {
            get => _includeGstr3B;
            set => SetProperty(ref _includeGstr3B, value);
        }

        public bool IncludeGstr1
        {
            get => _includeGstr1;
            set => SetProperty(ref _includeGstr1, value);
        }

        public bool IncludeGstr2A
        {
            get => _includeGstr2A;
            set => SetProperty(ref _includeGstr2A, value);
        }

        public bool IncludeGstr2B
        {
            get => _includeGstr2B;
            set => SetProperty(ref _includeGstr2B, value);
        }

        public bool Include3BVs1
        {
            get => _include3BVs1;
            set => SetProperty(ref _include3BVs1, value);
        }

        public bool Include3BVs2A
        {
            get => _include3BVs2A;
            set => SetProperty(ref _include3BVs2A, value);
        }

        public bool Include3BVs2B
        {
            get => _include3BVs2B;
            set => SetProperty(ref _include3BVs2B, value);
        }

        public bool CanInclude3B
        {
            get => _canInclude3B;
            set => SetProperty(ref _canInclude3B, value);
        }

        public bool CanInclude1
        {
            get => _canInclude1;
            set => SetProperty(ref _canInclude1, value);
        }

        public bool CanInclude2A
        {
            get => _canInclude2A;
            set => SetProperty(ref _canInclude2A, value);
        }

        public bool CanInclude2B
        {
            get => _canInclude2B;
            set => SetProperty(ref _canInclude2B, value);
        }

        public bool CanInclude3BVs1
        {
            get => _canInclude3BVs1;
            set => SetProperty(ref _canInclude3BVs1, value);
        }

        public bool CanInclude3BVs2A
        {
            get => _canInclude3BVs2A;
            set => SetProperty(ref _canInclude3BVs2A, value);
        }

        public bool CanInclude3BVs2B
        {
            get => _canInclude3BVs2B;
            set => SetProperty(ref _canInclude3BVs2B, value);
        }

        public bool IsBusy
        {
            get => _isBusy;
            set
            {
                if (SetProperty(ref _isBusy, value))
                {
                    OnPropertyChanged(nameof(CanGenerate));
                }
            }
        }

        public int ProgressPercentage
        {
            get => _progressPercentage;
            set => SetProperty(ref _progressPercentage, value);
        }

        public string ProgressStatusText
        {
            get => _progressStatusText;
            set => SetProperty(ref _progressStatusText, value);
        }

        public string ScanBadgeText
        {
            get => _scanBadgeText;
            set => SetProperty(ref _scanBadgeText, value);
        }

        public bool HasGenerationResult
        {
            get => _hasGenerationResult;
            set => SetProperty(ref _hasGenerationResult, value);
        }

        public string IncludedSheetsSummary
        {
            get => _includedSheetsSummary;
            set => SetProperty(ref _includedSheetsSummary, value);
        }

        public string SkippedSheetsSummary
        {
            get => _skippedSheetsSummary;
            set => SetProperty(ref _skippedSheetsSummary, value);
        }

        public string ErrorBannerText
        {
            get => _errorBannerText;
            set
            {
                if (SetProperty(ref _errorBannerText, value))
                {
                    OnPropertyChanged(nameof(HasErrorBanner));
                }
            }
        }

        public bool HasErrorBanner => !string.IsNullOrWhiteSpace(ErrorBannerText);
        public bool HasReconciliationAlerts => ReconciliationAlerts.Count > 0;

        #endregion

        #region Public Actions

        public async Task AddPathsAsync(IEnumerable<string> paths)
        {
            if (IsBusy) return;

            IsBusy = true;
            ErrorBannerText = string.Empty;
            ProgressPercentage = 20;
            ProgressStatusText = "Scanning JSON and ZIP files for GSTIN, Return Type & Month/Period...";

            try
            {
                _cts = new CancellationTokenSource();
                var newlyScanned = await _scanner.ScanPathsAsync(paths, _cts.Token);

                var existingKeys = new HashSet<string>(
                    _allScannedFiles.Select(f => $"{f.SourcePath}|{f.ZipEntryPath}"),
                    StringComparer.OrdinalIgnoreCase);

                foreach (var item in newlyScanned)
                {
                    string key = $"{item.SourcePath}|{item.ZipEntryPath}";
                    if (!existingKeys.Contains(key))
                    {
                        _allScannedFiles.Add(item);
                        existingKeys.Add(key);
                    }
                }

                if (string.IsNullOrWhiteSpace(OutputFolderPath) && _allScannedFiles.Count > 0)
                {
                    string? firstDir = Path.GetDirectoryName(_allScannedFiles[0].SourcePath);
                    if (!string.IsNullOrWhiteSpace(firstDir))
                    {
                        OutputFolderPath = Path.Combine(firstDir, "GST_Summary_Reports");
                    }
                }

                RebuildFinancialYearList();
                OnPropertyChanged(nameof(HasScannedFiles));
                OnPropertyChanged(nameof(CanGenerate));
                ProgressPercentage = 100;
                ProgressStatusText = "Scan complete. Ready to generate Summary & Reconciliation Excel.";
            }
            catch (Exception ex)
            {
                ErrorBannerText = $"Scan error: {ex.Message}";
                _logger.LogError("Summary Report scan failed", ex);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task SelectFilesAsync()
        {
            var files = _fileService.SelectJsonFiles();
            if (files != null && files.Length > 0)
            {
                await AddPathsAsync(files);
            }
        }

        private async Task SelectFolderAsync()
        {
            var folder = _fileService.SelectFolder();
            if (!string.IsNullOrWhiteSpace(folder))
            {
                await AddPathsAsync(new[] { folder });
            }
        }

        private void BrowseOutputFolder()
        {
            var folder = _fileService.SelectFolder();
            if (!string.IsNullOrWhiteSpace(folder))
            {
                OutputFolderPath = folder;
            }
        }

        public void ClearAll()
        {
            _cts?.Cancel();
            _allScannedFiles.Clear();
            AvailableFinancialYears.Clear();
            AvailableMonths.Clear();
            CoverageRows.Clear();
            CustomLabelItems.Clear();
            GeneratedFiles.Clear();
            ReconciliationAlerts.Clear();
            HasGenerationResult = false;
            ErrorBannerText = string.Empty;
            IncludedSheetsSummary = string.Empty;
            SkippedSheetsSummary = string.Empty;
            ProgressPercentage = 0;
            ProgressStatusText = "Ready to scan GST files";
            ScanBadgeText = "Drop or select 4 files (1 month's GSTR-1/3B/2A/2B), multiple months, or full annual files";
            CanInclude3B = false;
            CanInclude1 = false;
            CanInclude2A = false;
            CanInclude2B = false;
            CanInclude3BVs1 = false;
            CanInclude3BVs2A = false;
            CanInclude3BVs2B = false;
            OnPropertyChanged(nameof(HasScannedFiles));
            OnPropertyChanged(nameof(HasAvailableMonths));
            OnPropertyChanged(nameof(SelectedMonthsSummary));
            OnPropertyChanged(nameof(CanGenerate));
            OnPropertyChanged(nameof(HasReconciliationAlerts));
            OnPropertyChanged(nameof(HasGenerationResult));
            OnPropertyChanged(nameof(HasErrorBanner));
            CommandManager.InvalidateRequerySuggested();
        }

        private void RebuildFinancialYearList()
        {
            var fyGroups = _allScannedFiles
                .Where(f => !f.IsCorrupt && !string.IsNullOrWhiteSpace(f.FinancialYear))
                .GroupBy(f => f.FinancialYear)
                .OrderByDescending(g => g.Count())
                .ThenByDescending(g => g.Key)
                .ToList();

            AvailableFinancialYears.Clear();
            AvailableFinancialYears.Add(AllUploadedPeriodsOption);
            foreach (var g in fyGroups)
            {
                AvailableFinancialYears.Add(g.Key);
            }

            if (AvailableFinancialYears.Count > 1)
            {
                // Default to AllUploadedPeriodsOption so whether user gives 1 month or 12 months, all uploaded months are included
                _selectedFinancialYear = AllUploadedPeriodsOption;
                OnPropertyChanged(nameof(SelectedFinancialYear));
                RebuildAvailableMonthsForSelectedFy();
                RefreshCoverageForSelectedFy();
            }
            else
            {
                AvailableMonths.Clear();
                OnPropertyChanged(nameof(HasAvailableMonths));
                OnPropertyChanged(nameof(SelectedMonthsSummary));
                ScanBadgeText = $"Scanned {_allScannedFiles.Count} files — no valid GST return files detected.";
            }
        }

        private void RebuildAvailableMonthsForSelectedFy()
        {
            bool isAllOption = string.IsNullOrWhiteSpace(SelectedFinancialYear) ||
                               string.Equals(SelectedFinancialYear, AllUploadedPeriodsOption, StringComparison.OrdinalIgnoreCase);

            var fyFiles = _allScannedFiles
                .Where(f => !f.IsCorrupt && !f.IsDuplicate &&
                            (isAllOption || string.Equals(f.FinancialYear, SelectedFinancialYear, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            var prevUnselected = new HashSet<int>(
                AvailableMonths.Where(m => !m.IsSelected).Select(m => m.SortKey));

            AvailableMonths.Clear();
            foreach (var group in fyFiles.GroupBy(f => f.SortKey).OrderBy(g => g.Key))
            {
                var sample = group.First();
                var option = new GSTJsonToExcel.Models.SelectableMonthOption
                {
                    MonthLabel = sample.MonthLabel,
                    SortKey = group.Key,
                    FileCount = group.Count(),
                    OnSelectionChanged = _ =>
                    {
                        RefreshCoverageForSelectedFy();
                        OnPropertyChanged(nameof(SelectedMonthsSummary));
                        OnPropertyChanged(nameof(CanGenerate));
                    }
                };
                option.SetSelectedSilent(!prevUnselected.Contains(group.Key));
                AvailableMonths.Add(option);
            }

            OnPropertyChanged(nameof(HasAvailableMonths));
            OnPropertyChanged(nameof(SelectedMonthsSummary));
            OnPropertyChanged(nameof(CanGenerate));
        }

        private void SetAllMonthsSelection(bool isSelected)
        {
            foreach (var m in AvailableMonths)
            {
                m.SetSelectedSilent(isSelected);
            }
            RefreshCoverageForSelectedFy();
            OnPropertyChanged(nameof(SelectedMonthsSummary));
            OnPropertyChanged(nameof(CanGenerate));
        }

        private List<AnnualScannedFile> GetFilteredValidFiles()
        {
            bool isAllOption = string.IsNullOrWhiteSpace(SelectedFinancialYear) ||
                               string.Equals(SelectedFinancialYear, AllUploadedPeriodsOption, StringComparison.OrdinalIgnoreCase);

            var selectedSortKeys = AvailableMonths.Count > 0
                ? new HashSet<int>(AvailableMonths.Where(m => m.IsSelected).Select(m => m.SortKey))
                : null;

            return _allScannedFiles
                .Where(f => !f.IsCorrupt && !f.IsDuplicate &&
                            (isAllOption || string.Equals(f.FinancialYear, SelectedFinancialYear, StringComparison.OrdinalIgnoreCase)) &&
                            (selectedSortKeys == null || selectedSortKeys.Contains(f.SortKey)))
                .ToList();
        }

        private void RefreshCoverageForSelectedFy()
        {
            CoverageRows.Clear();
            if (_allScannedFiles.Count == 0) return;

            var validInSelection = GetFilteredValidFiles();
            int distinctMonthCount = Math.Max(1, validInSelection.Select(f => f.SortKey).Distinct().Count());
            int denominator = ForceFull12MonthColumns ? 12 : distinctMonthCount;

            int duplicateCount = _allScannedFiles.Count(f => f.IsDuplicate);
            int corruptCount = _allScannedFiles.Count(f => f.IsCorrupt);

            var gstins = validInSelection
                .Select(f => f.Gstin)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => g, StringComparer.OrdinalIgnoreCase)
                .ToList();

            string modeDesc = ForceFull12MonthColumns
                ? "12 Full FY Months Mode"
                : $"{distinctMonthCount} Uploaded Month(s) Dynamic Mode";

            ScanBadgeText = $"{validInSelection.Count} valid return files across {distinctMonthCount} month(s) & {gstins.Count} GSTIN(s) ({modeDesc})  •  {duplicateCount} duplicates skipped  •  {corruptCount} corrupt skipped";

            foreach (var g in gstins)
            {
                if (!CustomLabelItems.Any(x => string.Equals(x.Gstin, g, StringComparison.OrdinalIgnoreCase)))
                {
                    string state = Helpers.GstStateHelper.GetStateName(g);
                    CustomLabelItems.Add(new GstinCustomLabelItem
                    {
                        Gstin = g,
                        StateName = state,
                        CustomLabel = string.IsNullOrWhiteSpace(state) ? g : $"{state} Branch"
                    });
                }
            }

            AnnualReturnType[] types = { AnnualReturnType.GSTR3B, AnnualReturnType.GSTR1, AnnualReturnType.GSTR2A, AnnualReturnType.GSTR2B };
            foreach (var gstin in gstins)
            {
                string state = Helpers.GstStateHelper.GetStateName(gstin);
                foreach (var rt in types)
                {
                    var matching = validInSelection
                        .Where(f => string.Equals(f.Gstin, gstin, StringComparison.OrdinalIgnoreCase) && f.ReturnType == rt)
                        .ToList();

                    if (matching.Count == 0 && !validInSelection.Any(f => f.ReturnType == rt))
                    {
                        continue;
                    }

                    var row = new GstinMonthCoverageRow
                    {
                        Gstin = gstin,
                        StateName = state,
                        ActiveMonthsDenominator = denominator,
                        ActiveMonthsFoundCount = matching.Select(m => m.SortKey).Distinct().Count(),
                        ReturnTypeLabel = rt switch
                        {
                            AnnualReturnType.GSTR3B => "GSTR-3B",
                            AnnualReturnType.GSTR1 => "GSTR-1",
                            AnnualReturnType.GSTR2A => "GSTR-2A",
                            AnnualReturnType.GSTR2B => "GSTR-2B",
                            _ => "Unknown"
                        }
                    };

                    foreach (var mFile in matching)
                    {
                        if (mFile.FyMonthIndex >= 0 && mFile.FyMonthIndex < 12)
                        {
                            row.MonthFound[mFile.FyMonthIndex] = true;
                        }
                    }

                    CoverageRows.Add(row);
                }
            }

            CanInclude3B = validInSelection.Any(f => f.ReturnType == AnnualReturnType.GSTR3B);
            CanInclude1 = validInSelection.Any(f => f.ReturnType == AnnualReturnType.GSTR1);
            CanInclude2A = validInSelection.Any(f => f.ReturnType == AnnualReturnType.GSTR2A);
            CanInclude2B = validInSelection.Any(f => f.ReturnType == AnnualReturnType.GSTR2B);
            CanInclude3BVs1 = CanInclude3B && CanInclude1;
            CanInclude3BVs2A = CanInclude3B && CanInclude2A;
            CanInclude3BVs2B = CanInclude3B && CanInclude2B;
        }

        private async Task GenerateReportAsync()
        {
            if (!CanGenerate) return;

            IsBusy = true;
            HasGenerationResult = false;
            ErrorBannerText = string.Empty;
            GeneratedFiles.Clear();
            ReconciliationAlerts.Clear();
            ProgressPercentage = 15;
            ProgressStatusText = "Aggregating uploaded GST returns and building Summary & Reconciliation Excel...";

            try
            {
                _cts = new CancellationTokenSource();
                var settings = BuildCurrentSettings();
                SavePersistedSettings();

                var validFilesForReport = GetFilteredValidFiles();

                string outDir = string.IsNullOrWhiteSpace(OutputFolderPath)
                    ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "GST_Summary_Reports")
                    : OutputFolderPath;

                var includedSheets = new List<string>();
                var skippedReasons = new List<string>();
                var generatedPaths = new List<string>();
                AnnualWorkbookData? builtData = null;

                string targetFy = string.Equals(SelectedFinancialYear, AllUploadedPeriodsOption, StringComparison.OrdinalIgnoreCase)
                    ? (validFilesForReport.FirstOrDefault()?.FinancialYear ?? "2025-26")
                    : SelectedFinancialYear;

                await Task.Run(() =>
                {
                    builtData = _aggregator.BuildAnnualData(validFilesForReport, settings, targetFy);

                    string mainPath = _excelBuilder.BuildWorkbook(
                        builtData,
                        settings,
                        outDir,
                        includedSheets,
                        skippedReasons,
                        singleGstinFilter: null);
                    generatedPaths.Add(mainPath);

                    if (settings.SeparateWorkbookPerGstin && builtData.Gstins.Count > 1)
                    {
                        foreach (var gstin in builtData.Gstins)
                        {
                            var dummyIncluded = new List<string>();
                            var dummySkipped = new List<string>();
                            string branchPath = _excelBuilder.BuildWorkbook(
                                builtData,
                                settings,
                                outDir,
                                dummyIncluded,
                                dummySkipped,
                                singleGstinFilter: gstin);
                            generatedPaths.Add(branchPath);
                        }
                    }
                }, _cts.Token);

                foreach (var path in generatedPaths)
                {
                    GeneratedFiles.Add(path);
                }

                if (builtData != null)
                {
                    foreach (var alert in builtData.ReconciliationAlerts)
                    {
                        ReconciliationAlerts.Add(alert);
                    }
                }

                IncludedSheetsSummary = includedSheets.Count > 0
                    ? string.Join("  |  ", includedSheets.Distinct())
                    : "Overview";
                SkippedSheetsSummary = skippedReasons.Count > 0
                    ? string.Join(" • ", skippedReasons.Distinct())
                    : "None (All requested sheets generated)";

                HasGenerationResult = true;
                OnPropertyChanged(nameof(HasReconciliationAlerts));
                ProgressPercentage = 100;
                ProgressStatusText = $"Successfully generated {generatedPaths.Count} Annual / Summary Report workbook(s) (FY {builtData?.FinancialYear})!";
            }
            catch (Exception ex)
            {
                ErrorBannerText = $"Failed to generate Summary Report: {ex.Message}";
                _logger.LogError("Summary Report generation failed", ex);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private AnnualReportUserSettings BuildCurrentSettings()
        {
            var customMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in CustomLabelItems)
            {
                customMap[item.Gstin] = item.CustomLabel;
            }

            return new AnnualReportUserSettings
            {
                CompanyName = CompanyName,
                SelectedFinancialYear = SelectedFinancialYear,
                ForceFull12MonthColumns = ForceFull12MonthColumns,
                GstinDisplayMode = _gstinDisplayMode,
                CustomBranchLabels = customMap,
                IncludeOverview = IncludeOverview,
                IncludeGstr3B = IncludeGstr3B && CanInclude3B,
                IncludeGstr1 = IncludeGstr1 && CanInclude1,
                IncludeGstr2A = IncludeGstr2A && CanInclude2A,
                IncludeGstr2B = IncludeGstr2B && CanInclude2B,
                Include3BVs1 = Include3BVs1 && CanInclude3BVs1,
                Include3BVs2A = Include3BVs2A && CanInclude3BVs2A,
                Include3BVs2B = Include3BVs2B && CanInclude3BVs2B,
                SeparateWorkbookPerGstin = SeparateWorkbookPerGstin,
                OutputFolderPath = OutputFolderPath
            };
        }

        private static string GetSettingsFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string dir = Path.Combine(appData, "GSTJsonToExcel");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "annual_report_settings.json");
        }

        private void LoadPersistedSettings()
        {
            try
            {
                string path = GetSettingsFilePath();
                if (!File.Exists(path)) return;
                string json = File.ReadAllText(path);
                var saved = JsonSerializer.Deserialize<AnnualReportUserSettings>(json);
                if (saved != null)
                {
                    _companyName = saved.CompanyName ?? string.Empty;
                    _forceFull12MonthColumns = saved.ForceFull12MonthColumns;
                    _gstinDisplayMode = saved.GstinDisplayMode;
                    _separateWorkbookPerGstin = saved.SeparateWorkbookPerGstin;
                    _outputFolderPath = saved.OutputFolderPath ?? string.Empty;
                }
            }
            catch
            {
                // Ignore corrupt settings file
            }
        }

        private void SavePersistedSettings()
        {
            try
            {
                string path = GetSettingsFilePath();
                var settings = BuildCurrentSettings();
                string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(path, json);
            }
            catch
            {
                // Ignore IO errors when saving preferences
            }
        }

        private static void OpenFile(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                });
            }
            catch
            {
                // Ignore
            }
        }

        private static void ShowInFolder(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{path}\"",
                    UseShellExecute = true
                });
            }
            catch
            {
                // Ignore
            }
        }

        #endregion
    }
}
