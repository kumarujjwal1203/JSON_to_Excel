using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Features.AnnualReport.Models;
using GSTJsonToExcel.Features.AnnualReport.Services;
using Xunit;

namespace GSTJsonToExcel.Tests
{
    public class AnnualReportTests
    {
        private const string RealUserDataDir = @"C:\Users\ADMIN\Videos\DATA\Data for software\Data for software\Json";

        [Fact]
        public async Task SummaryReport_SingleMonth4Files_PopulatesThatMonthAndLeavesOtherMonthsEmpty()
        {
            if (!Directory.Exists(RealUserDataDir))
            {
                return;
            }

            var scanner = new AnnualReportScannerService();
            var scannedFiles = await scanner.ScanPathsAsync(new[] { RealUserDataDir });

            // Pick a single month ("012026" = Jan 2026)
            var singleMonthFiles = scannedFiles
                .Where(f => !f.IsCorrupt && !f.IsDuplicate && f.ReturnPeriod == "012026" && f.Gstin == "06ACWFS8659K1ZY")
                .ToList();

            Assert.NotEmpty(singleMonthFiles);

            var settings = new AnnualReportUserSettings
            {
                CompanyName = "Test_SPECTAL MANAGEMENT",
                SelectedFinancialYear = "2025-26",
                GstinDisplayMode = GstinDisplayMode.ShowGstinAndState,
                IncludeOverview = true,
                IncludeGstr3B = true,
                IncludeGstr1 = true,
                IncludeGstr2A = true,
                IncludeGstr2B = true,
                Include3BVs1 = true,
                Include3BVs2A = true,
                Include3BVs2B = true
            };

            var aggregator = new AnnualReportAggregatorService();
            var data = aggregator.BuildAnnualData(singleMonthFiles, settings, "2025-26");

            string tempOutDir = Path.Combine(Path.GetTempPath(), "GST_SingleMonth_Test_" + Guid.NewGuid().ToString("N")[..6]);
            try
            {
                var included = new List<string>();
                var skipped = new List<string>();
                var builder = new AnnualReportExcelBuilder();
                string excelPath = builder.BuildWorkbook(data, settings, tempOutDir, included, skipped);

                Assert.True(File.Exists(excelPath));
                using var wb = new XLWorkbook(excelPath);

                // Overview card check matching screenshot
                var wsOverview = wb.Worksheet("Overview");
                Assert.Equal("MAP & Associates", wsOverview.Cell(2, 2).GetString());
                Assert.Equal("Company Name", wsOverview.Cell(5, 2).GetString());
                Assert.Equal("Test_SPECTAL MANAGEMENT", wsOverview.Cell(5, 3).GetString());
                Assert.Equal("GST Returns Report", wsOverview.Cell(6, 3).GetString());
                Assert.Equal("2025-26", wsOverview.Cell(8, 3).GetString());
                Assert.Equal("Created by Octa GST", wsOverview.Cell(10, 2).GetString());

                // GSTR-3B should have 12 month columns (Apr 2025..Mar 2026) + Total = 16 columns
                var ws3B = wb.Worksheet("GSTR-3B");
                Assert.Equal("Company GSTIN", ws3B.Cell(1, 1).GetString());
                Assert.Equal("Section", ws3B.Cell(1, 2).GetString());
                Assert.Equal("Type", ws3B.Cell(1, 3).GetString());
                Assert.Equal("Apr 2025", ws3B.Cell(1, 4).GetString());
                Assert.Equal("Jan 2026", ws3B.Cell(1, 13).GetString()); // Slot 9 -> Col 13
                Assert.Equal("Mar 2026", ws3B.Cell(1, 15).GetString());
                Assert.Equal("Total", ws3B.Cell(1, 16).GetString());

                // Apr 2025 (Col 4) was NOT uploaded so it should be empty, while Jan 2026 (Col 13) WAS uploaded so it has value
                Assert.True(ws3B.Cell(2, 4).IsEmpty());
                Assert.False(ws3B.Cell(2, 13).IsEmpty());

                // GSTR-1 should have 24 month columns (Apr 2025..Mar 2027) + Total = 28 columns
                var ws1 = wb.Worksheet("GSTR-1");
                Assert.Equal("Company GSTIN", ws1.Cell(1, 1).GetString());
                Assert.Equal("Section", ws1.Cell(1, 2).GetString());
                Assert.Equal("Type", ws1.Cell(1, 3).GetString());
                Assert.Equal("Apr 2025", ws1.Cell(1, 4).GetString());
                Assert.Equal("Mar 2026", ws1.Cell(1, 15).GetString());
                Assert.Equal("Apr 2026", ws1.Cell(1, 16).GetString());
                Assert.Equal("Mar 2027", ws1.Cell(1, 27).GetString());
                Assert.Equal("Total", ws1.Cell(1, 28).GetString());
            }
            finally
            {
                if (Directory.Exists(tempOutDir))
                {
                    Directory.Delete(tempOutDir, true);
                }
            }
        }

        [Fact]
        public async Task AnnualReport_EndToEnd_WithRealUserData_GeneratesAll8ExactSheetsAndFormatting()
        {
            if (!Directory.Exists(RealUserDataDir))
            {
                return;
            }

            var scanner = new AnnualReportScannerService();
            var scannedFiles = await scanner.ScanPathsAsync(new[] { RealUserDataDir });

            Assert.NotEmpty(scannedFiles);

            string targetFy = "2025-26";
            var validForFy = scannedFiles
                .Where(f => !f.IsCorrupt && !f.IsDuplicate && f.FinancialYear == targetFy)
                .ToList();

            var settings = new AnnualReportUserSettings
            {
                CompanyName = "Test_SPECTAL MANAGEMENT",
                SelectedFinancialYear = targetFy,
                GstinDisplayMode = GstinDisplayMode.ShowGstinAndState,
                IncludeOverview = true,
                IncludeGstr3B = true,
                IncludeGstr1 = true,
                IncludeGstr2A = true,
                IncludeGstr2B = true,
                Include3BVs1 = true,
                Include3BVs2A = true,
                Include3BVs2B = true
            };

            var aggregator = new AnnualReportAggregatorService();
            var data = aggregator.BuildAnnualData(validForFy, settings, targetFy);

            string tempOutDir = Path.Combine(Path.GetTempPath(), "GST_Annual_Report_Test_" + Guid.NewGuid().ToString("N")[..6]);
            try
            {
                var included = new List<string>();
                var skipped = new List<string>();
                var builder = new AnnualReportExcelBuilder();
                string excelPath = builder.BuildWorkbook(data, settings, tempOutDir, included, skipped);

                Assert.True(File.Exists(excelPath));

                using var wb = new XLWorkbook(excelPath);
                string[] expectedOrder =
                {
                    "Overview",
                    "GSTR-3B",
                    "GSTR-1",
                    "GSTR-2A",
                    "GSTR-2B",
                    "GSTR-3B vs GSTR-1",
                    "GSTR-3B vs GSTR-2A",
                    "GSTR-3B vs GSTR-2B"
                };

                var actualOrder = wb.Worksheets.Select(w => w.Name).ToArray();
                Assert.Equal(expectedOrder, actualOrder);

                // Verify 24-month headers on GSTR-2A and GSTR-2B
                foreach (string sName in new[] { "GSTR-1", "GSTR-2A", "GSTR-2B" })
                {
                    var ws = wb.Worksheet(sName);
                    Assert.Equal("Company GSTIN", ws.Cell(1, 1).GetString());
                    Assert.Equal("Section", ws.Cell(1, 2).GetString());
                    Assert.Equal("Type", ws.Cell(1, 3).GetString());
                    Assert.Equal("Apr 2025", ws.Cell(1, 4).GetString());
                    Assert.Equal("Mar 2026", ws.Cell(1, 15).GetString());
                    Assert.Equal("Apr 2026", ws.Cell(1, 16).GetString());
                    Assert.Equal("Mar 2027", ws.Cell(1, 27).GetString());
                    Assert.Equal("Total", ws.Cell(1, 28).GetString());
                }

                // Verify 12-month headers on GSTR-3B and 3 reconciliation sheets
                foreach (string sName in new[] { "GSTR-3B", "GSTR-3B vs GSTR-1", "GSTR-3B vs GSTR-2A", "GSTR-3B vs GSTR-2B" })
                {
                    var ws = wb.Worksheet(sName);
                    Assert.Equal("Company GSTIN", ws.Cell(1, 1).GetString());
                    Assert.Equal("Section", ws.Cell(1, 2).GetString());
                    Assert.Equal("Type", ws.Cell(1, 3).GetString());
                    Assert.Equal("Apr 2025", ws.Cell(1, 4).GetString());
                    Assert.Equal("Mar 2026", ws.Cell(1, 15).GetString());
                    Assert.Equal("Total", ws.Cell(1, 16).GetString());
                }
            }
            finally
            {
                if (Directory.Exists(tempOutDir))
                {
                    Directory.Delete(tempOutDir, true);
                }
            }
        }
    }
}
