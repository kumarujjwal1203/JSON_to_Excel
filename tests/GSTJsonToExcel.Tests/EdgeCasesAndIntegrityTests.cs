using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Services.Implementations;
using Xunit;

namespace GSTJsonToExcel.Tests
{
    public class EdgeCasesAndIntegrityTests
    {
        private readonly LoggingService _logger = new();

        [Fact]
        public async Task Convert_EdgeCases_PreservesNulls_EmptyArrays_AndFutureFields()
        {
            // Arrange
            string jsonPath = Path.Combine(AppContext.BaseDirectory, "TestData", "sample_edge_cases.json");
            string tempOut = Path.Combine(Path.GetTempPath(), $"EdgeCases_Test_{Guid.NewGuid():N}.xlsx");

            var parser = new JsonParserService(_logger);
            var exporter = new ExcelExportService(_logger);
            var verifier = new IntegrityCheckService(_logger);

            try
            {
                // Act
                var parseResult = await parser.ParseAndFlattenAsync(jsonPath);
                var exportResult = await exporter.ExportToExcelAsync(parseResult, tempOut);
                var integrityReport = await verifier.VerifyIntegrityAsync(parseResult, tempOut);

                Assert.True(exportResult.Success);
                Assert.True(integrityReport.Passed);

                using var wb = new XLWorkbook(tempOut);

                // 1. Verify unknown future field is preserved
                bool foundFutureField = false;
                foreach (var sheet in wb.Worksheets)
                {
                    foreach (var cell in sheet.CellsUsed())
                    {
                        if (cell.GetString() == "FutureGSTStandard_v3.2")
                        {
                            foundFutureField = true;
                        }
                    }
                }
                Assert.True(foundFutureField, "Future field 'FutureGSTStandard_v3.2' was dropped!");

                // 2. Verify null values are retained (not silently dropped)
                bool foundNullField = false;
                foreach (var sheet in wb.Worksheets)
                {
                    foreach (var cell in sheet.CellsUsed())
                    {
                        if (cell.GetString() == "(null)")
                        {
                            foundNullField = true;
                        }
                    }
                }
                Assert.True(foundNullField, "Null value representation was dropped!");

                // 3. Verify empty array is recorded in Empty_Sections
                var emptySheet = wb.Worksheets.FirstOrDefault(s => s.Name == "Empty_Sections");
                Assert.NotNull(emptySheet);
                bool foundEmptySection = emptySheet.CellsUsed().Any(c => c.GetString().Contains("empty_array_test"));
                Assert.True(foundEmptySection, "Empty array 'empty_array_test' was not audited in Empty_Sections!");

                // 4. Verify negative number
                bool foundNegative = false;
                foreach (var sheet in wb.Worksheets)
                {
                    foreach (var cell in sheet.CellsUsed())
                    {
                        if (cell.TryGetValue(out double d) && Math.Abs(d - (-4500.5)) < 0.01)
                        {
                            foundNegative = true;
                        }
                    }
                }
                Assert.True(foundNegative, "Negative adjustment -4500.5 was not preserved!");
            }
            finally
            {
                if (File.Exists(tempOut)) try { File.Delete(tempOut); } catch { }
            }
        }
    }
}
