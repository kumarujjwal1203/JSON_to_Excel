using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Services.Implementations;
using Xunit;

namespace GSTJsonToExcel.Tests
{
    public class SampleGstr3bConversionTests
    {
        private readonly LoggingService _logger = new();

        [Fact]
        public async Task Convert_PromptSampleGstr3b_ZeroDataLoss_IntegrityPassed()
        {
            // Arrange
            string jsonPath = Path.Combine(AppContext.BaseDirectory, "TestData", "sample_gstr3b.json");
            Assert.True(File.Exists(jsonPath), $"Test data file not found at: {jsonPath}");

            string tempOutputPath = Path.Combine(Path.GetTempPath(), $"GSTR3B_Test_{Guid.NewGuid():N}.xlsx");

            var parser = new JsonParserService(_logger);
            var exporter = new ExcelExportService(_logger);
            var verifier = new IntegrityCheckService(_logger);

            try
            {
                // Act 1: Parse and Flatten
                var parseResult = await parser.ParseAndFlattenAsync(jsonPath);

                Assert.NotNull(parseResult);
                Assert.True(parseResult.TotalLeafCount > 100, $"Expected > 100 leaf values, got {parseResult.TotalLeafCount}");
                Assert.Contains("$.gstin", parseResult.AllLeafNodes.Keys);
                Assert.Equal("24AAACO4007A7Z5", parseResult.AllLeafNodes["$.gstin"].RawString);
                Assert.Equal("052026", parseResult.AllLeafNodes["$.ret_period"].RawString);

                // Act 2: Export to Excel
                var exportResult = await exporter.ExportToExcelAsync(parseResult, tempOutputPath);

                Assert.True(exportResult.Success, exportResult.ErrorMessage);
                Assert.True(File.Exists(tempOutputPath));

                // Act 3: Verify Integrity
                var integrityReport = await verifier.VerifyIntegrityAsync(parseResult, tempOutputPath);

                Assert.True(integrityReport.Passed, $"Integrity check failed: {string.Join("; ", integrityReport.Warnings)}");
                Assert.Equal(parseResult.TotalLeafCount, integrityReport.MatchedLeafValues);

                // Act 4: Inspect Excel Workbook with ClosedXML
                using var wb = new XLWorkbook(tempOutputPath);

                // 1) Verify Summary sheet exists
                var summarySheet = wb.Worksheet("Summary");
                Assert.NotNull(summarySheet);
                Assert.Contains("PASSED", summarySheet.Cell(12, 2).GetString());

                // 2) Verify General_Info sheet
                var generalSheet = wb.Worksheets.FirstOrDefault(s => s.Name == "General_Info");
                Assert.NotNull(generalSheet);
                var gstinCell = generalSheet.Row(2).Cells().FirstOrDefault(c => c.GetString() == "24AAACO4007A7Z5");
                Assert.NotNull(gstinCell);

                // 3) Verify sup_details sheet has categories
                var supSheet = wb.Worksheets.FirstOrDefault(s => s.Name == "sup_details");
                Assert.NotNull(supSheet);
                Assert.True(supSheet.RowsUsed().Count() >= 5); // osup_det, osup_zero, osup_nil_exmp, isup_rev, osup_nongst

                // 4) Verify itc_avl array sheet
                var itcSheet = wb.Worksheets.FirstOrDefault(s => s.Name.Contains("itc_avl"));
                Assert.NotNull(itcSheet);
                // In prompt, ISD has iamt = 9919.73
                var isdRow = itcSheet.RowsUsed().FirstOrDefault(r => r.Cells().Any(c => c.GetString() == "ISD"));
                Assert.NotNull(isdRow);
                var isdVal = isdRow.Cells().FirstOrDefault(c => c.TryGetValue(out double d) && Math.Abs(d - 9919.73) < 0.001);
                Assert.NotNull(isdVal);

                // 5) Verify eco_dtls with 58917516.44
                var ecoSheet = wb.Worksheets.FirstOrDefault(s => s.Name.Contains("eco_dtls"));
                Assert.NotNull(ecoSheet);
                var ecoValCell = ecoSheet.CellsUsed().FirstOrDefault(c => c.TryGetValue(out double d) && Math.Abs(d - 58917516.44) < 0.01);
                Assert.NotNull(ecoValCell);

                // 6) Verify All_Data_Index has all leaves
                var indexSheet = wb.Worksheet("All_Data_Index");
                Assert.NotNull(indexSheet);
                Assert.Equal(parseResult.TotalLeafCount, indexSheet.RowsUsed().Count() - 1);
            }
            finally
            {
                if (File.Exists(tempOutputPath))
                {
                    try { File.Delete(tempOutputPath); } catch { }
                }
            }
        }
    }
}
