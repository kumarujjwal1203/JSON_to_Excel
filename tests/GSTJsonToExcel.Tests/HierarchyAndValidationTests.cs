using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Services.Implementations;
using Xunit;

namespace GSTJsonToExcel.Tests
{
    public class HierarchyAndValidationTests
    {
        private readonly LoggingService _logger = new();
        private readonly FileService _fileService = new();

        [Fact]
        public async Task Convert_Gstr1_ParentKeysInheritedByChildRecords()
        {
            // Arrange
            string jsonPath = Path.Combine(AppContext.BaseDirectory, "TestData", "sample_gstr1.json");
            string tempOut = Path.Combine(Path.GetTempPath(), $"Gstr1_Hierarchy_{Guid.NewGuid():N}.xlsx");

            var parser = new JsonParserService(_logger);
            var exporter = new ExcelExportService(_logger);
            var verifier = new IntegrityCheckService(_logger);

            try
            {
                // Act
                var parseResult = await parser.ParseAndFlattenAsync(jsonPath);
                var exportResult = await exporter.ExportToExcelAsync(parseResult, tempOut);
                var integrity = await verifier.VerifyIntegrityAsync(parseResult, tempOut);

                Assert.True(exportResult.Success);
                Assert.True(integrity.Passed);

                using var wb = new XLWorkbook(tempOut);

                // Child sheets should inherit parent identifiers like parent_gstin, parent_fp, or parent_ctin
                var b2bInvSheet = wb.Worksheets.FirstOrDefault(s => s.Name.Contains("b2b_inv") || s.Name.Contains("inv"));
                Assert.NotNull(b2bInvSheet);

                // Ensure parent_gstin or parent_ctin column exists in child table
                bool hasParentKeyCol = b2bInvSheet.Row(1).Cells().Any(c => c.GetString().StartsWith("parent_"));
                Assert.True(hasParentKeyCol, "Child invoice sheet does not contain parent linkage keys!");
            }
            finally
            {
                if (File.Exists(tempOut)) try { File.Delete(tempOut); } catch { }
            }
        }

        [Fact]
        public void ValidateJsonFile_HandlesMalformedAndMissingFilesGracefully()
        {
            // 1. Non-existent file
            var nonExistent = _fileService.ValidateJsonFile("C:\\NonExistent_File_12345.json");
            Assert.False(nonExistent.IsValid);
            Assert.Contains("does not exist", nonExistent.ErrorMessage, StringComparison.OrdinalIgnoreCase);

            // 2. Empty file
            string emptyFile = Path.Combine(Path.GetTempPath(), $"Empty_{Guid.NewGuid():N}.json");
            File.WriteAllText(emptyFile, "");
            try
            {
                var emptyResult = _fileService.ValidateJsonFile(emptyFile);
                Assert.False(emptyResult.IsValid);
                Assert.Contains("empty", emptyResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                if (File.Exists(emptyFile)) File.Delete(emptyFile);
            }

            // 3. Corrupted / Malformed JSON
            string corruptFile = Path.Combine(Path.GetTempPath(), $"Corrupt_{Guid.NewGuid():N}.json");
            File.WriteAllText(corruptFile, "{ \"gstin\": \"24AAACO4007A7Z5\", unclosed... ");
            try
            {
                var corruptResult = _fileService.ValidateJsonFile(corruptFile);
                Assert.False(corruptResult.IsValid);
                Assert.Contains("Invalid JSON", corruptResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                if (File.Exists(corruptFile)) File.Delete(corruptFile);
            }
        }
    }
}
