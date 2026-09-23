using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Services.Implementations;
using Xunit;

namespace GSTJsonToExcel.Tests
{
    public class LeadingZerosAndPrecisionTests
    {
        private readonly LoggingService _logger = new();

        [Fact]
        public async Task Convert_PreservesLeadingZeros_NeverConvertsToInteger()
        {
            // Arrange
            string jsonPath = Path.Combine(AppContext.BaseDirectory, "TestData", "sample_gstr1.json");
            string tempOut = Path.Combine(Path.GetTempPath(), $"LeadingZeros_Test_{Guid.NewGuid():N}.xlsx");

            var parser = new JsonParserService(_logger);
            var exporter = new ExcelExportService(_logger);

            try
            {
                // Act
                var parseResult = await parser.ParseAndFlattenAsync(jsonPath);
                var exportResult = await exporter.ExportToExcelAsync(parseResult, tempOut);

                Assert.True(exportResult.Success);

                using var wb = new XLWorkbook(tempOut);

                // 1. Verify inum "00001234" is preserved with leading zeros in both section and index sheets
                bool foundLeadingZeroInvoice = false;
                foreach (var sheet in wb.Worksheets)
                {
                    foreach (var cell in sheet.CellsUsed())
                    {
                        if (cell.GetString() == "00001234")
                        {
                            foundLeadingZeroInvoice = true;
                            // Verify cell format or text data type
                            Assert.Equal("00001234", cell.GetString());
                            Assert.NotEqual("1234", cell.GetString());
                        }
                    }
                }
                Assert.True(foundLeadingZeroInvoice, "Invoice number '00001234' with leading zeros was not preserved!");

                // 2. Verify HSN "008471" leading zeros
                bool foundHsn = false;
                foreach (var sheet in wb.Worksheets)
                {
                    foreach (var cell in sheet.CellsUsed())
                    {
                        if (cell.GetString() == "008471")
                        {
                            foundHsn = true;
                            Assert.NotEqual("8471", cell.GetString());
                        }
                    }
                }
                Assert.True(foundHsn, "HSN code '008471' with leading zeros was not preserved!");

                // 3. Verify high precision decimal "123456.7891"
                bool foundPrecisionVal = false;
                foreach (var sheet in wb.Worksheets)
                {
                    foreach (var cell in sheet.CellsUsed())
                    {
                        if (cell.TryGetValue(out decimal dec) && Math.Abs(dec - 123456.7891m) < 0.00001m)
                        {
                            foundPrecisionVal = true;
                        }
                        else if (cell.GetString().Contains("123456.7891"))
                        {
                            foundPrecisionVal = true;
                        }
                    }
                }
                Assert.True(foundPrecisionVal, "Decimal '123456.7891' precision was lost or truncated!");
            }
            finally
            {
                if (File.Exists(tempOut)) try { File.Delete(tempOut); } catch { }
            }
        }
    }
}
