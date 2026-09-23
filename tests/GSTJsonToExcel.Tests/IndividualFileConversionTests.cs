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
    public class IndividualFileConversionTests
    {
        private readonly LoggingService _logger = new();

        [Fact]
        public async Task ConvertBatch_IndividualMode_GeneratesSeparateExcelPerJsonFile_NeverLumpsIntoOne()
        {
            var parser = new JsonParserService(_logger);
            var exporter = new ExcelExportService(_logger);
            var integrity = new IntegrityCheckService(_logger);
            var octaBuilder = new OctaGstBuilderService(_logger);
            var batchService = new BatchConversionService(parser, exporter, integrity, _logger, octaBuilder);

            string tempInputDir = Path.Combine(Path.GetTempPath(), $"IndivIn_{Guid.NewGuid():N}");
            string tempOutputDir = Path.Combine(Path.GetTempPath(), $"IndivOut_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempInputDir);

            try
            {
                // Create 3 separate R2B JSON files with distinct invoice numbers
                string r2b_1 = Path.Combine(tempInputDir, "R2B_April.json");
                File.WriteAllText(r2b_1, @"{
                    ""gstin"": ""07ACWFS8659K2ZV"",
                    ""fp"": ""042025"",
                    ""data"": {
                        ""docdata"": {
                            ""b2b"": [{ ""ctin"": ""27AAPFU0939F1ZV"", ""inv"": [{ ""inum"": ""INV-APR-001"", ""val"": 1000, ""itms"": [{ ""itm_det"": { ""txval"": 1000, ""rt"": 18 } }] }] }]
                        }
                    }
                }");

                string r2b_2 = Path.Combine(tempInputDir, "R2B_May.json");
                File.WriteAllText(r2b_2, @"{
                    ""gstin"": ""07ACWFS8659K2ZV"",
                    ""fp"": ""052025"",
                    ""data"": {
                        ""docdata"": {
                            ""b2b"": [{ ""ctin"": ""27AAPFU0939F1ZV"", ""inv"": [{ ""inum"": ""INV-MAY-002"", ""val"": 2000, ""itms"": [{ ""itm_det"": { ""txval"": 2000, ""rt"": 18 } }] }] }]
                        }
                    }
                }");

                string r2b_3 = Path.Combine(tempInputDir, "R2B_June.json");
                File.WriteAllText(r2b_3, @"{
                    ""gstin"": ""07ACWFS8659K2ZV"",
                    ""fp"": ""062025"",
                    ""data"": {
                        ""docdata"": {
                            ""b2b"": [{ ""ctin"": ""27AAPFU0939F1ZV"", ""inv"": [{ ""inum"": ""INV-JUN-003"", ""val"": 3000, ""itms"": [{ ""itm_det"": { ""txval"": 3000, ""rt"": 18 } }] }] }]
                        }
                    }
                }");

                var summary = new BatchScanSummary();
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r2b_1,
                    FileName = "R2B_April.json",
                    FileType = GstFileType.R2B,
                    Status = FileValidationStatus.Valid,
                    IsSelected = true
                });
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r2b_2,
                    FileName = "R2B_May.json",
                    FileType = GstFileType.R2B,
                    Status = FileValidationStatus.Valid,
                    IsSelected = true
                });
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r2b_3,
                    FileName = "R2B_June.json",
                    FileType = GstFileType.R2B,
                    Status = FileValidationStatus.Valid,
                    IsSelected = true
                });

                // Act: Convert in IndividualFiles mode (Default)
                var result = await batchService.ConvertBatchAsync(
                    summary,
                    tempOutputDir,
                    BatchConversionMode.IndividualFiles);

                // Assert
                Assert.True(result.Success);
                Assert.Equal(3, result.GeneratedFiles.Count);

                string out1 = Path.Combine(tempOutputDir, "R2B_April.xlsx");
                string out2 = Path.Combine(tempOutputDir, "R2B_May.xlsx");
                string out3 = Path.Combine(tempOutputDir, "R2B_June.xlsx");
                string mergedFile = Path.Combine(tempOutputDir, "R2B.xlsx");

                // Each individual file MUST exist
                Assert.True(File.Exists(out1), "R2B_April.xlsx was not generated!");
                Assert.True(File.Exists(out2), "R2B_May.xlsx was not generated!");
                Assert.True(File.Exists(out3), "R2B_June.xlsx was not generated!");

                // The combined file R2B.xlsx MUST NOT exist in individual mode
                Assert.False(File.Exists(mergedFile), "Files must NOT be lumped into a single R2B.xlsx in individual mode!");

                // Check that each file contains ONLY its own data
                using (var wb1 = new XLWorkbook(out1))
                {
                    var ws = wb1.Worksheet("Purchase");
                    Assert.Contains(ws.CellsUsed(), c => c.GetString() == "INV-APR-001");
                    Assert.DoesNotContain(ws.CellsUsed(), c => c.GetString() == "INV-MAY-002");
                    Assert.DoesNotContain(ws.CellsUsed(), c => c.GetString() == "INV-JUN-003");
                }

                using (var wb2 = new XLWorkbook(out2))
                {
                    var ws = wb2.Worksheet("Purchase");
                    Assert.DoesNotContain(ws.CellsUsed(), c => c.GetString() == "INV-APR-001");
                    Assert.Contains(ws.CellsUsed(), c => c.GetString() == "INV-MAY-002");
                    Assert.DoesNotContain(ws.CellsUsed(), c => c.GetString() == "INV-JUN-003");
                }

                using (var wb3 = new XLWorkbook(out3))
                {
                    var ws = wb3.Worksheet("Purchase");
                    Assert.DoesNotContain(ws.CellsUsed(), c => c.GetString() == "INV-APR-001");
                    Assert.DoesNotContain(ws.CellsUsed(), c => c.GetString() == "INV-MAY-002");
                    Assert.Contains(ws.CellsUsed(), c => c.GetString() == "INV-JUN-003");
                }
            }
            finally
            {
                if (Directory.Exists(tempInputDir)) Directory.Delete(tempInputDir, true);
                if (Directory.Exists(tempOutputDir)) Directory.Delete(tempOutputDir, true);
            }
        }

        [Fact]
        public async Task ConvertBatch_IndividualMode_UncheckedItemsAreSkipped()
        {
            var parser = new JsonParserService(_logger);
            var exporter = new ExcelExportService(_logger);
            var integrity = new IntegrityCheckService(_logger);
            var octaBuilder = new OctaGstBuilderService(_logger);
            var batchService = new BatchConversionService(parser, exporter, integrity, _logger, octaBuilder);

            string tempInputDir = Path.Combine(Path.GetTempPath(), $"IndivSkip_{Guid.NewGuid():N}");
            string tempOutputDir = Path.Combine(Path.GetTempPath(), $"IndivSkipOut_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempInputDir);

            try
            {
                string r1_1 = Path.Combine(tempInputDir, "R1_001.json");
                File.WriteAllText(r1_1, "{ \"gstin\": \"07ACWFS8659K2ZV\", \"fp\": \"042025\" }");

                string r1_2 = Path.Combine(tempInputDir, "R1_002.json");
                File.WriteAllText(r1_2, "{ \"gstin\": \"07ACWFS8659K2ZV\", \"fp\": \"052025\" }");

                var summary = new BatchScanSummary();
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r1_1,
                    FileName = "R1_001.json",
                    FileType = GstFileType.R1,
                    Status = FileValidationStatus.Valid,
                    IsSelected = true // Checked
                });
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r1_2,
                    FileName = "R1_002.json",
                    FileType = GstFileType.R1,
                    Status = FileValidationStatus.Valid,
                    IsSelected = false // Unchecked by user in UI
                });

                var result = await batchService.ConvertBatchAsync(
                    summary,
                    tempOutputDir,
                    BatchConversionMode.IndividualFiles);

                Assert.True(result.Success);
                Assert.Single(result.GeneratedFiles);
                Assert.True(File.Exists(Path.Combine(tempOutputDir, "R1_001.xlsx")));
                Assert.False(File.Exists(Path.Combine(tempOutputDir, "R1_002.xlsx")));
            }
            finally
            {
                if (Directory.Exists(tempInputDir)) Directory.Delete(tempInputDir, true);
                if (Directory.Exists(tempOutputDir)) Directory.Delete(tempOutputDir, true);
            }
        }

        [Fact]
        public async Task ConvertBatch_RealUserFolder_ConvertsAllFilesAccurately()
        {
            string userJsonDir = @"C:\Users\ADMIN\Videos\DATA\Data for software\Data for software\Json";
            if (!Directory.Exists(userJsonDir)) return;

            var logger = new LoggingService();
            var classifier = new GstClassifierService(logger);
            var duplicateDetector = new DuplicateDetectorService(logger);
            var fileService = new FileService();
            var scanner = new FileScannerService(classifier, duplicateDetector, fileService, logger);

            var jsonFiles = Directory.GetFiles(userJsonDir, "*.json");
            var summary = await scanner.ScanPathsAsync(jsonFiles);

            Assert.Equal(24, summary.CandidateGstFiles.Count(f => f.Status == FileValidationStatus.Valid));

            var parser = new JsonParserService(logger);
            var exporter = new ExcelExportService(logger);
            var integrity = new IntegrityCheckService(logger);
            var octaBuilder = new OctaGstBuilderService(logger);
            var batchService = new BatchConversionService(parser, exporter, integrity, logger, octaBuilder);

            string outDir = Path.Combine(userJsonDir, "GST_Converted");
            Directory.CreateDirectory(outDir);

            var result = await batchService.ConvertBatchAsync(summary, outDir, BatchConversionMode.IndividualFiles);

            Assert.True(result.Success, $"Conversion failed: {result.FailedCount} errors");
            Assert.Equal(24, result.TotalJsonFilesProcessed);

            // Verify each generated Excel has rows > 1 in Purchase sheet
            foreach (var outItem in result.GeneratedFiles)
            {
                using var wb = new XLWorkbook(outItem.OutputFilePath);
                var purchSheet = wb.Worksheet("Purchase");
                Assert.NotNull(purchSheet);
                int lastRow = purchSheet.LastRowUsed()?.RowNumber() ?? 0;
                Assert.True(lastRow > 1, $"Excel {outItem.FileName} has 0 data rows in Purchase!");
            }
        }
    }
}
