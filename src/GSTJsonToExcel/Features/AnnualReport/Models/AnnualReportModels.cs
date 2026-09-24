using System;
using System.Collections.Generic;
using System.Linq;
using GSTJsonToExcel.ViewModels;

namespace GSTJsonToExcel.Features.AnnualReport.Models
{
    public enum AnnualReturnType
    {
        Unknown = 0,
        GSTR3B = 1,
        GSTR1 = 2,
        GSTR2A = 3,
        GSTR2B = 4
    }

    public enum GstinDisplayMode
    {
        ShowGstinAndState = 0,
        ShowStateOnly = 1,
        ShowCustomBranchName = 2
    }

    public class AnnualScannedFile
    {
        public string SourcePath { get; set; } = string.Empty;
        public string? ZipEntryPath { get; set; }
        public string DisplayFileName { get; set; } = string.Empty;
        public AnnualReturnType ReturnType { get; set; } = AnnualReturnType.Unknown;
        public string Gstin { get; set; } = string.Empty;
        public string StateName { get; set; } = string.Empty;
        public string ReturnPeriod { get; set; } = string.Empty; // MMYYYY e.g. 042025
        public int MonthNumber { get; set; } // 1..12
        public int YearNumber { get; set; } // e.g. 2025
        public int SortKey => (YearNumber * 100) + MonthNumber;
        public int FyMonthIndex { get; set; } = -1; // 0 = Apr .. 11 = Mar
        public string FinancialYear { get; set; } = string.Empty; // e.g. "2025-26"
        public DateTime LastModified { get; set; }
        public bool IsDuplicate { get; set; }
        public bool IsOutsideSelectedFy { get; set; }
        public bool IsCorrupt { get; set; }
        public string StatusReason { get; set; } = "Ready";

        private static readonly string[] MonthNames =
        {
            "", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        };

        public string MonthLabel => MonthNumber is >= 1 and <= 12
            ? $"{MonthNames[MonthNumber]} {YearNumber}"
            : ReturnPeriod;

        public string ReturnTypeShort => ReturnType switch
        {
            AnnualReturnType.GSTR3B => "GSTR-3B",
            AnnualReturnType.GSTR1 => "GSTR-1",
            AnnualReturnType.GSTR2A => "GSTR-2A",
            AnnualReturnType.GSTR2B => "GSTR-2B",
            _ => "Unknown"
        };
    }

    public class GstinCustomLabelItem : ViewModelBase
    {
        private string _customLabel = string.Empty;
        public string Gstin { get; set; } = string.Empty;
        public string StateName { get; set; } = string.Empty;

        public string CustomLabel
        {
            get => _customLabel;
            set => SetProperty(ref _customLabel, value);
        }
    }

    public class GstinMonthCoverageRow
    {
        public string Gstin { get; set; } = string.Empty;
        public string StateName { get; set; } = string.Empty;
        public string ReturnTypeLabel { get; set; } = string.Empty;
        public bool[] MonthFound { get; set; } = new bool[12]; // 0=Apr .. 11=Mar
        public int ActiveMonthsDenominator { get; set; } = 12;
        public int ActiveMonthsFoundCount { get; set; }

        public bool Apr => MonthFound[0];
        public bool May => MonthFound[1];
        public bool Jun => MonthFound[2];
        public bool Jul => MonthFound[3];
        public bool Aug => MonthFound[4];
        public bool Sep => MonthFound[5];
        public bool Oct => MonthFound[6];
        public bool Nov => MonthFound[7];
        public bool Dec => MonthFound[8];
        public bool Jan => MonthFound[9];
        public bool Feb => MonthFound[10];
        public bool Mar => MonthFound[11];

        public int FoundMonthsCount => ActiveMonthsFoundCount > 0 ? ActiveMonthsFoundCount : MonthFound.Count(m => m);
        public string SummaryBadge => $"{FoundMonthsCount}/12";
    }

    public class AnnualSheetRow
    {
        public AnnualSheetRow(int monthColumnsCount = 12)
        {
            int count = Math.Max(1, monthColumnsCount);
            MonthlyValues = new decimal[count];
            HasMonthData = new bool[count];
        }

        public string Gstin { get; set; } = string.Empty;
        public string GstinDisplay { get; set; } = string.Empty; // Col A: Company GSTIN
        public string Section { get; set; } = string.Empty;      // Col B: Section
        public string Description { get; set; } = string.Empty;  // Col C: Type (Taxable, IGST, CGST, SGST, Cess, etc.)
        public int SortOrder { get; set; }
        public int SubOrder { get; set; }
        public bool IsSectionSummaryRow { get; set; } // #BDD7EE background + bold
        public bool IsDifferenceRow { get; set; }     // #D7D7D7 background + bold

        // Length 12 for GSTR-3B & Reconciliations (Apr YYYY .. Mar YYYY+1)
        // Length 24 for GSTR-1, GSTR-2A, GSTR-2B (Apr YYYY .. Mar YYYY+1 + Apr YYYY+1 .. Mar YYYY+2)
        public decimal[] MonthlyValues { get; set; }
        public bool[] HasMonthData { get; set; }

        public decimal TotalValue => MonthlyValues.Sum();

        public bool HasAnyNonZeroData => MonthlyValues.Any(v => v != 0m);
    }

    public class AnnualOverviewGstinSummary
    {
        public string Gstin { get; set; } = string.Empty;
        public string GstinDisplay { get; set; } = string.Empty;
        public string StateName { get; set; } = string.Empty;
        public int TotalActiveMonths { get; set; } = 12;
        public int R3BMonthsCount { get; set; }
        public int R1MonthsCount { get; set; }
        public int R2AMonthsCount { get; set; }
        public int R2BMonthsCount { get; set; }

        public decimal OutwardTaxable3B { get; set; }
        public decimal OutwardTax3B { get; set; }
        public decimal OutwardTaxable1 { get; set; }
        public decimal OutwardTax1 { get; set; }
        public decimal NetItc3B { get; set; }
        public decimal Itc2A { get; set; }
        public decimal Itc2B { get; set; }
        public decimal CashPaid3B { get; set; }
        public decimal ItcPaid3B { get; set; }
    }

    public class AnnualReconciliationAlert
    {
        public string GstinDisplay { get; set; } = string.Empty;
        public string ComparisonSheet { get; set; } = string.Empty;
        public string MetricName { get; set; } = string.Empty;
        public decimal DifferenceAmount { get; set; }
        public string FormattedDifference => $"₹ {DifferenceAmount:N2}";
    }

    public class AnnualWorkbookData
    {
        public string FirmHeaderTitle { get; set; } = "MAP & Associates";
        public string CompanyName { get; set; } = string.Empty;
        public string FinancialYear { get; set; } = string.Empty; // e.g. "2025-26"
        public string PeriodRangeLabel { get; set; } = string.Empty;
        public int StartYear { get; set; } // e.g. 2025
        public int EndYear { get; set; }   // e.g. 2026
        public int NextEndYear => EndYear + 1; // e.g. 2027 (for 24-month sheets Apr 2025..Mar 2027)
        public List<string> Gstins { get; set; } = new();
        public Dictionary<string, string> GstinDisplayMap { get; set; } = new();

        public bool Has3BData { get; set; }
        public bool Has1Data { get; set; }
        public bool Has2AData { get; set; }
        public bool Has2BData { get; set; }

        public List<AnnualOverviewGstinSummary> OverviewSummaries { get; set; } = new();
        public List<AnnualSheetRow> Gstr3BRows { get; set; } = new();
        public List<AnnualSheetRow> Gstr1Rows { get; set; } = new();
        public List<AnnualSheetRow> Gstr2ARows { get; set; } = new();
        public List<AnnualSheetRow> Gstr2BRows { get; set; } = new();
        public List<AnnualSheetRow> Gstr3BVs1Rows { get; set; } = new();
        public List<AnnualSheetRow> Gstr3BVs2ARows { get; set; } = new();
        public List<AnnualSheetRow> Gstr3BVs2BRows { get; set; } = new();
        public List<AnnualReconciliationAlert> ReconciliationAlerts { get; set; } = new();
    }

    public class AnnualReportUserSettings
    {
        public string CompanyName { get; set; } = string.Empty;
        public string SelectedFinancialYear { get; set; } = string.Empty;
        public bool ForceFull12MonthColumns { get; set; } = true;
        public GstinDisplayMode GstinDisplayMode { get; set; } = GstinDisplayMode.ShowGstinAndState;
        public Dictionary<string, string> CustomBranchLabels { get; set; } = new();
        public bool IncludeOverview { get; set; } = true;
        public bool IncludeGstr3B { get; set; } = true;
        public bool IncludeGstr1 { get; set; } = true;
        public bool IncludeGstr2A { get; set; } = true;
        public bool IncludeGstr2B { get; set; } = true;
        public bool Include3BVs1 { get; set; } = true;
        public bool Include3BVs2A { get; set; } = true;
        public bool Include3BVs2B { get; set; } = true;
        public bool SeparateWorkbookPerGstin { get; set; }
        public string OutputFolderPath { get; set; } = string.Empty;
    }
}
