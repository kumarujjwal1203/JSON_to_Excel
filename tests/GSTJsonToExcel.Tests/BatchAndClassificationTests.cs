using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Implementations;
using Xunit;

namespace GSTJsonToExcel.Tests
{
    public class BatchAndClassificationTests
    {
        private readonly LoggingService _logger = new();
        private readonly FileService _fileService = new();

        [Theory]
        [InlineData("R1_001.json", GstFileType.R1)]
        [InlineData("GSTR1_2026_09.json", GstFileType.R1)]
        [InlineData("R1_September.json", GstFileType.R1)]
        [InlineData("R3A_001.json", GstFileType.R3A)]
        [InlineData("GSTR3B_052026.json", GstFileType.R3A)]
        [InlineData("R3A_Sep_Data.json", GstFileType.R3A)]
        [InlineData("R2A_001.json", GstFileType.R2A)]
        [InlineData("GSTR2A_2026.json", GstFileType.R2A)]
        [InlineData("R2B_001.json", GstFileType.R2B)]
        [InlineData("GSTR2B_Sep.json", GstFileType.R2B)]
        [InlineData("random_file.json", GstFileType.Unknown)]
        public void ClassifyByFileName_AccuratelyIdentifiesTypes(string fileName, GstFileType expectedType)
        {
            var classifier = new GstClassifierService(_logger);
            var (type, _) = classifier.ClassifyByFileName(fileName);
            Assert.Equal(expectedType, type);
        }

        [Fact]
        public async Task ClassifyFileAsync_FallsBackToContentInspection_ForAmbiguousNames()
        {
            var classifier = new GstClassifierService(_logger);
            string tempDir = Path.Combine(Path.GetTempPath(), $"GstClassify_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                // Ambiguous file name, but contains R3A/R3B structure
                string r3File = Path.Combine(tempDir, "data_001.json");
                File.WriteAllText(r3File, "{ \"gstin\": \"24AAACO4007A7Z5\", \"sup_details\": { \"txval\": 100 } }");

                var (type, reason) = await classifier.ClassifyFileAsync(r3File);
                Assert.Equal(GstFileType.R3A, type);
                Assert.Contains("content inspection", reason);

                // Ambiguous file name, but contains R1 structure
                string r1File = Path.Combine(tempDir, "export_sept.json");
                File.WriteAllText(r1File, "{ \"gstin\": \"09ABCDE1234F1Z5\", \"b2b\": [], \"hsn\": [] }");

                var (typeR1, reasonR1) = await classifier.ClassifyFileAsync(r1File);
                Assert.Equal(GstFileType.R1, typeR1);
                Assert.Contains("content inspection", reasonR1);

                // Ambiguous file name with unrelated content
                string unknownFile = Path.Combine(tempDir, "other_config.json");
                File.WriteAllText(unknownFile, "{ \"theme\": \"dark\", \"timeout\": 30 }");

                var (typeUnk, reasonUnk) = await classifier.ClassifyFileAsync(unknownFile);
                Assert.Equal(GstFileType.Unknown, typeUnk);
                Assert.Contains("Ambiguous", reasonUnk);
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task DuplicateDetector_IdentifiesIdenticalContent_ByHash()
        {
            var duplicateDetector = new DuplicateDetectorService(_logger);
            string tempDir = Path.Combine(Path.GetTempPath(), $"GstDup_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                string originalPath = Path.Combine(tempDir, "R1_September.json");
                string copyPath = Path.Combine(tempDir, "R1_September_Copy.json");
                string content = "{ \"gstin\": \"24AAACO4007A7Z5\", \"b2b\": [{ \"inum\": \"001\" }] }";

                File.WriteAllText(originalPath, content);
                File.WriteAllText(copyPath, content); // Exact same content

                var items = new List<ScannedFileItem>
                {
                    new() { FilePath = originalPath, FileName = "R1_September.json", Status = FileValidationStatus.Valid },
                    new() { FilePath = copyPath, FileName = "R1_September_Copy.json", Status = FileValidationStatus.Valid }
                };

                await duplicateDetector.DetectDuplicatesAsync(items);

                Assert.Equal(FileValidationStatus.Valid, items[0].Status);
                Assert.Equal(FileValidationStatus.Duplicate, items[1].Status);
                Assert.Equal("R1_September.json", items[1].DuplicateOf);
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task FileScanner_IgnoresNonJsonFiles_LeavesUserFilesUntouched()
        {
            var classifier = new GstClassifierService(_logger);
            var duplicateDetector = new DuplicateDetectorService(_logger);
            var scanner = new FileScannerService(classifier, duplicateDetector, _fileService, _logger);

            string tempDir = Path.Combine(Path.GetTempPath(), $"GstScan_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                // Create mixed files
                File.WriteAllText(Path.Combine(tempDir, "R1_001.json"), "{ \"gstin\": \"24AAACO4007A7Z5\" }");
                File.WriteAllText(Path.Combine(tempDir, "R3A_001.json"), "{ \"gstin\": \"24AAACO4007A7Z5\" }");
                File.WriteAllText(Path.Combine(tempDir, "sample.pdf"), "binary pdf content");
                File.WriteAllText(Path.Combine(tempDir, "document.xml"), "<xml>data</xml>");
                File.WriteAllText(Path.Combine(tempDir, "image.png"), "binary png content");
                File.WriteAllText(Path.Combine(tempDir, "notes.txt"), "some notes");

                var summary = await scanner.ScanPathsAsync(new[] { tempDir });

                Assert.Equal(6, summary.TotalFilesFound);
                Assert.Equal(2, summary.CandidateGstFiles.Count);
                Assert.Equal(4, summary.IgnoredCount);

                // Verify user files are 100% UNTOUCHED
                Assert.True(File.Exists(Path.Combine(tempDir, "sample.pdf")));
                Assert.True(File.Exists(Path.Combine(tempDir, "document.xml")));
                Assert.True(File.Exists(Path.Combine(tempDir, "image.png")));
                Assert.True(File.Exists(Path.Combine(tempDir, "notes.txt")));
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task BatchConversion_GeneratesSeparateExcelPerGstType_NeverMixesData()
        {
            var parser = new JsonParserService(_logger);
            var exporter = new ExcelExportService(_logger);
            var integrity = new IntegrityCheckService(_logger);
            var batchService = new BatchConversionService(parser, exporter, integrity, _logger);

            string tempInputDir = Path.Combine(Path.GetTempPath(), $"BatchIn_{Guid.NewGuid():N}");
            string tempOutputDir = Path.Combine(Path.GetTempPath(), $"BatchOut_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempInputDir);

            try
            {
                // Create 2 R1 files
                string r1_1 = Path.Combine(tempInputDir, "R1_001.json");
                File.WriteAllText(r1_1, "{ \"gstin\": \"09ABCDE1234F1Z5\", \"b2b\": [{ \"inum\": \"INV001\", \"val\": 1000 }] }");

                string r1_2 = Path.Combine(tempInputDir, "R1_002.json");
                File.WriteAllText(r1_2, "{ \"gstin\": \"09ABCDE1234F1Z5\", \"b2b\": [{ \"inum\": \"INV002\", \"val\": 2000 }] }");

                // Create 1 R3A file
                string r3_1 = Path.Combine(tempInputDir, "R3A_001.json");
                File.WriteAllText(r3_1, "{ \"gstin\": \"24AAACO4007A7Z5\", \"sup_details\": { \"osup_det\": { \"txval\": 5000 } } }");

                var summary = new BatchScanSummary();
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r1_1,
                    FileName = "R1_001.json",
                    FileType = GstFileType.R1,
                    Status = FileValidationStatus.Valid
                });
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r1_2,
                    FileName = "R1_002.json",
                    FileType = GstFileType.R1,
                    Status = FileValidationStatus.Valid
                });
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r3_1,
                    FileName = "R3A_001.json",
                    FileType = GstFileType.R3A,
                    Status = FileValidationStatus.Valid
                });

                // Act: Convert Batch in MergedByType mode
                var result = await batchService.ConvertBatchAsync(summary, tempOutputDir, BatchConversionMode.MergedByType);

                // Assert
                Assert.True(result.Success);
                Assert.Equal(2, result.GeneratedFiles.Count); // R1 and R3A

                string r1OutPath = Path.Combine(tempOutputDir, "R1.xlsx");
                string r3OutPath = Path.Combine(tempOutputDir, "R3A.xlsx");
                string r2aOutPath = Path.Combine(tempOutputDir, "R2A.xlsx");
                string r2bOutPath = Path.Combine(tempOutputDir, "R2B.xlsx");

                // Separate Excel files exist for present types
                Assert.True(File.Exists(r1OutPath), "R1.xlsx was not generated!");
                Assert.True(File.Exists(r3OutPath), "R3A.xlsx was not generated!");

                // Types with 0 files are NOT generated
                Assert.False(File.Exists(r2aOutPath), "R2A.xlsx should not exist when 0 R2A files were provided!");
                Assert.False(File.Exists(r2bOutPath), "R2B.xlsx should not exist when 0 R2B files were provided!");

                // Inspect R1.xlsx with ClosedXML: verify both R1_001.json and R1_002.json data are included
                using var wbR1 = new XLWorkbook(r1OutPath);
                var b2bSheet = wbR1.Worksheets.FirstOrDefault(s => s.Name.Contains("b2b"));
                Assert.NotNull(b2bSheet);

                // Both invoices INV001 and INV002 must be present in R1.xlsx
                bool foundInv1 = b2bSheet.CellsUsed().Any(c => c.GetString() == "INV001");
                bool foundInv2 = b2bSheet.CellsUsed().Any(c => c.GetString() == "INV002");
                Assert.True(foundInv1 && foundInv2, "R1.xlsx did not aggregate all R1 source records!");
            }
            finally
            {
                if (Directory.Exists(tempInputDir)) Directory.Delete(tempInputDir, true);
                if (Directory.Exists(tempOutputDir)) Directory.Delete(tempOutputDir, true);
            }
        }
    }
}
