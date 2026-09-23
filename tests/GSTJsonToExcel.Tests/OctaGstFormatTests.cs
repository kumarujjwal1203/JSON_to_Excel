using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Helpers;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Implementations;
using Xunit;

namespace GSTJsonToExcel.Tests
{
    public class OctaGstFormatTests
    {
        private readonly LoggingService _logger = new();

        [Theory]
        [InlineData("07ACWFS8659K2ZV", "Delhi", "07ACWFS8659K2ZV (Delhi)")]
        [InlineData("06ACWFS8659K1ZY", "Haryana", "06ACWFS8659K1ZY (Haryana)")]
        [InlineData("27AAPFU0939F1ZV", "Maharashtra", "27AAPFU0939F1ZV (Maharashtra)")]
        [InlineData("24AAACO4007A7Z5", "Gujarat", "24AAACO4007A7Z5 (Gujarat)")]
        [InlineData("33AAACG1234A1Z1", "Tamil Nadu", "33AAACG1234A1Z1 (Tamil Nadu)")]
        public void GstStateHelper_FormatsGstinCorrectly(string gstin, string expectedState, string expectedFormatted)
        {
            Assert.Equal(expectedState, GstStateHelper.GetStateName(gstin));
            Assert.Equal(expectedFormatted, GstStateHelper.FormatGstinWithState(gstin));
        }

        [Fact]
        public async Task OctaGstBuilder_R1_GeneratesExactSheetsAndColumns()
        {
            var builder = new OctaGstBuilderService(_logger);
            string tempDir = Path.Combine(Path.GetTempPath(), $"OctaR1_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                string r1JsonPath = Path.Combine(tempDir, "R1_Test.json");
                string jsonContent = @"{
                    ""gstin"": ""07ACWFS8659K2ZV"",
                    ""fp"": ""042025"",
                    ""cname"": ""Test_SPECTAL MANAGEMENT"",
                    ""b2b"": [
                        {
                            ""ctin"": ""06ACWFS8659K1ZY"",
                            ""cname"": ""Alpha Corp"",
                            ""inv"": [
                                {
                                    ""inum"": ""INV-001"",
                                    ""idt"": ""15-04-2025"",
                                    ""val"": 11800.00,
                                    ""pos"": ""06"",
                                    ""rchrg"": ""N"",
                                    ""itms"": [
                                        {
                                            ""num"": 1,
                                            ""itm_det"": {
                                                ""rt"": 18.0,
                                                ""txval"": 10000.00,
                                                ""iamt"": 1800.00,
                                                ""camt"": 0.0,
                                                ""samt"": 0.0,
                                                ""csamt"": 0.0
                                            }
                                        }
                                    ]
                                }
                            ]
                        }
                    ],
                    ""b2cs"": [
                        {
                            ""pos"": ""07"",
                            ""txval"": 5000.0,
                            ""rt"": 18.0,
                            ""camt"": 450.0,
                            ""samt"": 450.0
                        }
                    ],
                    ""hsn"": [
                        {
                            ""hsn_sc"": ""998311"",
                            ""desc"": ""IT Consulting"",
                            ""uqc"": ""NOS"",
                            ""qty"": 1.0,
                            ""val"": 11800.0,
                            ""txval"": 10000.0,
                            ""rt"": 18.0,
                            ""iamt"": 1800.0
                        }
                    ],
                    ""doc_issue"": {
                        ""doc_det"": [
                            {
                                ""doc_num"": 1,
                                ""docs"": [
                                    { ""from"": ""INV-001"", ""to"": ""INV-010"", ""totnum"": 10, ""canc"": 0 }
                                ]
                            }
                        ]
                    }
                }";
                await File.WriteAllTextAsync(r1JsonPath, jsonContent);

                string outExcel = Path.Combine(tempDir, "R1.xlsx");
                var files = new List<ScannedFileItem>
                {
                    new() { FilePath = r1JsonPath, FileName = "R1_Test.json", FileType = GstFileType.R1, Status = FileValidationStatus.Valid }
                };

                var (success, totalRecords, error) = await builder.BuildOctaWorkbookAsync(GstFileType.R1, files, outExcel);

                Assert.True(success, error);
                Assert.True(File.Exists(outExcel));

                using var wb = new XLWorkbook(outExcel);

                // 1. Verify Sheet Names
                var expectedSheets = new[] { "Overview", "Sales", "Sales Summary", "SalesHSN", "Disclosed", "All_Data_Index" };
                foreach (var sheetName in expectedSheets)
                {
                    Assert.NotNull(wb.Worksheet(sheetName));
                }

                // 2. Verify Overview sheet contents
                var wsOverview = wb.Worksheet("Overview");
                Assert.Equal("MAP & Associates", wsOverview.Cell("A1").GetString());
                Assert.Equal("Company Name", wsOverview.Cell("A3").GetString());
                Assert.Equal("Test_SPECTAL MANAGEMENT", wsOverview.Cell("B3").GetString());
                Assert.Equal("Contents", wsOverview.Cell("A4").GetString());
                Assert.Equal("GSTR-1/1A Data", wsOverview.Cell("B4").GetString());
                Assert.Equal("Company GSTIN", wsOverview.Cell("A5").GetString());
                Assert.Contains("07ACWFS8659K2ZV (Delhi)", wsOverview.Cell("B5").GetString());

                // 3. Verify Sales Header Columns
                string[] expectedSalesCols = {
                    "Company GSTIN", "Tax Period", "Doc Type", "Sale Type", "Doc No", "Doc Date",
                    "Customer GSTIN", "Customer Name", "Place of Supply", "Shipping Bill Date",
                    "Shipping Bill No", "Port Code", "Reference Doc No", "Reference Doc Date",
                    "Reverse Charge", "Doc Value", "Item Taxable Value", "GST Rate", "IGST",
                    "CGST", "SGST", "Cess", "Is Amendment", "Original Doc No", "Original Doc Date",
                    "Uploaded By", "Source", "IRN Date", "IRN", "GSTR-1A"
                };
                var wsSales = wb.Worksheet("Sales");
                for (int i = 0; i < expectedSalesCols.Length; i++)
                {
                    Assert.Equal(expectedSalesCols[i], wsSales.Cell(1, i + 1).GetString().Trim());
                }

                // Verify Sales Data row
                Assert.Equal("07ACWFS8659K2ZV", wsSales.Cell(2, 1).GetString());
                Assert.Equal("INV-001", wsSales.Cell(2, 5).GetString());
                Assert.Equal("06ACWFS8659K1ZY", wsSales.Cell(2, 7).GetString());
                Assert.Equal("Alpha Corp", wsSales.Cell(2, 8).GetString());

                // 4. Verify Sales Summary Columns
                string[] expectedSummaryCols = {
                    "Company GSTIN", "Tax Period", "Summary Type", "Description", "Place of Supply",
                    "Taxable Value", "GST Rate", "IGST", "CGST", "SGST", "Cess", "Ecommerce GSTIN",
                    "Original Period", "Original Ecommerce GSTIN", "GSTR-1A"
                };
                var wsSummary = wb.Worksheet("Sales Summary");
                for (int i = 0; i < expectedSummaryCols.Length; i++)
                {
                    Assert.Equal(expectedSummaryCols[i], wsSummary.Cell(1, i + 1).GetString().Trim());
                }

                // 5. Verify SalesHSN Columns
                string[] expectedHsnCols = {
                    "Company GSTIN", "Tax Period", "Summary Type", "HSN Code", "HSN Description",
                    "UQC", "Qty", "Total Value", "Taxable Value", "GST Rate", "IGST", "CGST",
                    "SGST", "Cess", "GSTR-1A"
                };
                var wsHsn = wb.Worksheet("SalesHSN");
                for (int i = 0; i < expectedHsnCols.Length; i++)
                {
                    Assert.Equal(expectedHsnCols[i], wsHsn.Cell(1, i + 1).GetString().Trim());
                }

                // 6. Verify Disclosed Columns
                string[] expectedDiscCols = {
                    "Company GSTIN", "Tax Period", "Doc Type", "From", "To", "Total", "Cancelled", "GSTR-1A"
                };
                var wsDisc = wb.Worksheet("Disclosed");
                for (int i = 0; i < expectedDiscCols.Length; i++)
                {
                    Assert.Equal(expectedDiscCols[i], wsDisc.Cell(1, i + 1).GetString().Trim());
                }

                // 7. Verify All_Data_Index has leaves
                var wsIndex = wb.Worksheet("All_Data_Index");
                Assert.True(wsIndex.RowCount() > 1, "All_Data_Index must contain index entries!");
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task OctaGstBuilder_R2A_GeneratesExactSheetsAndColumns()
        {
            var builder = new OctaGstBuilderService(_logger);
            string tempDir = Path.Combine(Path.GetTempPath(), $"OctaR2A_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                string r2aJsonPath = Path.Combine(tempDir, "R2A_Test.json");
                string jsonContent = @"{
                    ""gstin"": ""07ACWFS8659K2ZV"",
                    ""fp"": ""042025"",
                    ""b2b"": [
                        {
                            ""ctin"": ""27AAPFU0939F1ZV"",
                            ""cname"": ""Beta Traders"",
                            ""cfs"": ""Y"",
                            ""inv"": [
                                {
                                    ""inum"": ""PUR-101"",
                                    ""idt"": ""10-04-2025"",
                                    ""val"": 5900.0,
                                    ""pos"": ""07"",
                                    ""rchrg"": ""N"",
                                    ""itms"": [
                                        {
                                            ""num"": 1,
                                            ""itm_det"": {
                                                ""rt"": 18.0,
                                                ""txval"": 5000.0,
                                                ""iamt"": 900.0,
                                                ""camt"": 0.0,
                                                ""samt"": 0.0,
                                                ""csamt"": 0.0
                                            }
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }";
                await File.WriteAllTextAsync(r2aJsonPath, jsonContent);

                string outExcel = Path.Combine(tempDir, "R2A.xlsx");
                var files = new List<ScannedFileItem>
                {
                    new() { FilePath = r2aJsonPath, FileName = "R2A_Test.json", FileType = GstFileType.R2A, Status = FileValidationStatus.Valid }
                };

                var (success, _, error) = await builder.BuildOctaWorkbookAsync(GstFileType.R2A, files, outExcel);

                Assert.True(success, error);

                using var wb = new XLWorkbook(outExcel);

                // Sheets: Overview, Purchase, ISD, TDS, TCS, All_Data_Index
                var expectedSheets = new[] { "Overview", "Purchase", "ISD", "TDS", "TCS", "All_Data_Index" };
                foreach (var s in expectedSheets)
                {
                    Assert.NotNull(wb.Worksheet(s));
                }

                // Verify Overview Contents
                var wsOverview = wb.Worksheet("Overview");
                Assert.Equal("GSTR-2A", wsOverview.Cell("B4").GetString());

                // Verify Purchase Columns (33 cols)
                string[] expectedPurchCols = {
                    "Company GSTIN", "Tax Period", "Doc Type", "Purchase Type", "Doc No", "Doc Date",
                    "Supplier GSTIN", "Supplier Name", "Supplier State", "Place of Supply", "Port Code",
                    "Reference Doc No", "Reference Doc Date", "Reverse Charge", "Doc Value",
                    "Item Taxable Value", "GST Rate", "IGST", "CGST", "SGST", "Cess", "Is Amendment",
                    "Original Doc No", "Original Doc Date", "GSTR-1 Status", "GSTR-1 Filing Date",
                    "GSTR-3B Status", "Cancellation Date", "GSTR-9 (8A) ITC Available", "Uploaded By",
                    "Source", "IRN Date", "IRN"
                };
                var wsPurch = wb.Worksheet("Purchase");
                for (int i = 0; i < expectedPurchCols.Length; i++)
                {
                    Assert.Equal(expectedPurchCols[i], wsPurch.Cell(1, i + 1).GetString().Trim());
                }

                // Verify Supplier State populated as Maharashtra
                Assert.Equal("PUR-101", wsPurch.Cell(2, 5).GetString());
                Assert.Equal("Maharashtra", wsPurch.Cell(2, 9).GetString());

                // Verify TCS Columns (11 cols)
                string[] expectedTcsCols = {
                    "Company GSTIN", "Tax Period", "Collector GSTIN", "Collector Name", "Collector Tax Period",
                    "Gross Value of Supplies", "Value of Supplies Returned", "Net Amount Liable to TCS", "IGST", "CGST", "SGST"
                };
                var wsTcs = wb.Worksheet("TCS");
                for (int i = 0; i < expectedTcsCols.Length; i++)
                {
                    Assert.Equal(expectedTcsCols[i], wsTcs.Cell(1, i + 1).GetString().Trim());
                }
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task OctaGstBuilder_R2B_GeneratesExactSheetsAndColumns()
        {
            var builder = new OctaGstBuilderService(_logger);
            string tempDir = Path.Combine(Path.GetTempPath(), $"OctaR2B_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                string r2bJsonPath = Path.Combine(tempDir, "R2B_Test.json");
                string jsonContent = @"{
                    ""gstin"": ""07ACWFS8659K2ZV"",
                    ""fp"": ""052025"",
                    ""data"": {
                        ""docdata"": {
                            ""b2b"": [
                                {
                                    ""ctin"": ""24AAACO4007A7Z5"",
                                    ""cname"": ""Gamma Industries"",
                                    ""inv"": [
                                        {
                                            ""inum"": ""INV-999"",
                                            ""idt"": ""05-05-2025"",
                                            ""val"": 2360.0,
                                            ""pos"": ""24"",
                                            ""rev"": ""N"",
                                            ""itcavl"": ""Y"",
                                            ""itms"": [
                                                {
                                                    ""num"": 1,
                                                    ""itm_det"": {
                                                        ""rt"": 18.0,
                                                        ""txval"": 2000.0,
                                                        ""iamt"": 360.0,
                                                        ""camt"": 0.0,
                                                        ""samt"": 0.0,
                                                        ""csamt"": 0.0
                                                    }
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ]
                        }
                    }
                }";
                await File.WriteAllTextAsync(r2bJsonPath, jsonContent);

                string outExcel = Path.Combine(tempDir, "R2B.xlsx");
                var files = new List<ScannedFileItem>
                {
                    new() { FilePath = r2bJsonPath, FileName = "R2B_Test.json", FileType = GstFileType.R2B, Status = FileValidationStatus.Valid }
                };

                var (success, _, error) = await builder.BuildOctaWorkbookAsync(GstFileType.R2B, files, outExcel);

                Assert.True(success, error);

                using var wb = new XLWorkbook(outExcel);

                // Sheets: Overview, Purchase, ISD, All_Data_Index
                var expectedSheets = new[] { "Overview", "Purchase", "ISD", "All_Data_Index" };
                foreach (var s in expectedSheets)
                {
                    Assert.NotNull(wb.Worksheet(s));
                }

                // Verify Overview Contents
                var wsOverview = wb.Worksheet("Overview");
                Assert.Equal("GSTR-2B Data", wsOverview.Cell("B4").GetString());

                // Verify Purchase Columns (39 cols)
                string[] expectedPurchCols = {
                    "Company GSTIN", "Tax Period", "Doc Type", "Purchase Type", "Doc No", "Doc Date",
                    "Supplier GSTIN", "Supplier Name", "Supplier State", "Place of Supply", "Port Code",
                    "Reverse Charge", "Doc Value", "Item Taxable Value", "GST Rate", "IGST",
                    "CGST", "SGST", "Cess", "Is Amendment", "Original Doc No", "Original Doc Date",
                    "IMS Action", "GSTR-1 Filing Period", "GSTR-1 Filing Date", "ITC Eligible",
                    "ITC Ineligible Reason Type", "GSTR-9 (8A) ITC Available", "ICEGATE Reference Date",
                    "ICEGATE Received Date", "Uploaded By", "Source", "IRN Date", "IRN",
                    "IMS ITC-IGST", "IMS ITC-CGST", "IMS ITC-SGST", "IMS ITC-Cess", "IMS Remarks"
                };
                var wsPurch = wb.Worksheet("Purchase");
                for (int i = 0; i < expectedPurchCols.Length; i++)
                {
                    Assert.Equal(expectedPurchCols[i], wsPurch.Cell(1, i + 1).GetString().Trim());
                }

                // Verify Supplier State populated as Gujarat
                Assert.Equal("INV-999", wsPurch.Cell(2, 5).GetString());
                Assert.Equal("Gujarat", wsPurch.Cell(2, 9).GetString());
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task BatchConversionService_WithOctaBuilder_GeneratesOctaForR1_AndStandardForR3A()
        {
            var parser = new JsonParserService(_logger);
            var exporter = new ExcelExportService(_logger);
            var integrity = new IntegrityCheckService(_logger);
            var octaBuilder = new OctaGstBuilderService(_logger);

            var batchService = new BatchConversionService(parser, exporter, integrity, _logger, octaBuilder);

            string tempInputDir = Path.Combine(Path.GetTempPath(), $"OctaBatchIn_{Guid.NewGuid():N}");
            string tempOutputDir = Path.Combine(Path.GetTempPath(), $"OctaBatchOut_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempInputDir);

            try
            {
                // R1 File
                string r1Path = Path.Combine(tempInputDir, "R1_001.json");
                File.WriteAllText(r1Path, "{ \"gstin\": \"07ACWFS8659K2ZV\", \"fp\": \"062025\", \"b2b\": [{ \"ctin\": \"06ACWFS8659K1ZY\", \"inv\": [{ \"inum\": \"INV-777\", \"val\": 500, \"itms\": [{ \"itm_det\": { \"txval\": 500, \"rt\": 18 } }] }] }] }");

                // R3A File
                string r3Path = Path.Combine(tempInputDir, "R3A_001.json");
                File.WriteAllText(r3Path, "{ \"gstin\": \"07ACWFS8659K2ZV\", \"fp\": \"062025\", \"sup_details\": { \"osup_det\": { \"txval\": 9999 } } }");

                var summary = new BatchScanSummary();
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r1Path,
                    FileName = "R1_001.json",
                    FileType = GstFileType.R1,
                    Status = FileValidationStatus.Valid
                });
                summary.CandidateGstFiles.Add(new ScannedFileItem
                {
                    FilePath = r3Path,
                    FileName = "R3A_001.json",
                    FileType = GstFileType.R3A,
                    Status = FileValidationStatus.Valid
                });

                var result = await batchService.ConvertBatchAsync(summary, tempOutputDir, BatchConversionMode.MergedByType);

                Assert.True(result.Success);
                Assert.Equal(2, result.GeneratedFiles.Count);

                string r1Out = Path.Combine(tempOutputDir, "R1.xlsx");
                string r3Out = Path.Combine(tempOutputDir, "R3A.xlsx");

                Assert.True(File.Exists(r1Out));
                Assert.True(File.Exists(r3Out));

                // R1.xlsx should have Octa sheets: Overview, Sales, etc.
                using var wbR1 = new XLWorkbook(r1Out);
                Assert.NotNull(wbR1.Worksheet("Overview"));
                Assert.NotNull(wbR1.Worksheet("Sales"));
                Assert.NotNull(wbR1.Worksheet("Sales Summary"));

                var salesSheet = wbR1.Worksheet("Sales");
                Assert.Contains(salesSheet.CellsUsed(), c => c.GetString() == "INV-777");

                // R3A.xlsx should have generic flattened sheets + All_Data_Index
                using var wbR3 = new XLWorkbook(r3Out);
                Assert.NotNull(wbR3.Worksheet("All_Data_Index"));
                Assert.Contains(wbR3.Worksheets, s => s.Name.Contains("sup_details") || s.Name.Contains("Table"));
            }
            finally
            {
                if (Directory.Exists(tempInputDir)) Directory.Delete(tempInputDir, true);
                if (Directory.Exists(tempOutputDir)) Directory.Delete(tempOutputDir, true);
            }
        }

        [Fact]
        public async Task OctaGstBuilder_RealR2BFile_MatchesGovSummaryAccurately()
        {
            string realR2BPath = @"C:\Users\ADMIN\Videos\DATA\Data for software\Data for software\Json\06ACWFS8659K1ZY_R2B_012026_20260921.json";
            if (!File.Exists(realR2BPath)) return;

            var builder = new OctaGstBuilderService(_logger);
            string tempDir = Path.Combine(Path.GetTempPath(), $"OctaRealR2B_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                string outExcel = Path.Combine(tempDir, "Real_R2B.xlsx");
                var files = new List<ScannedFileItem>
                {
                    new() { FilePath = realR2BPath, FileName = Path.GetFileName(realR2BPath), FileType = GstFileType.R2B, Status = FileValidationStatus.Valid }
                };

                var (success, count, error) = await builder.BuildOctaWorkbookAsync(GstFileType.R2B, files, outExcel);
                Assert.True(success, error);
                Assert.Equal(37, count);

                using var wb = new XLWorkbook(outExcel);
                var wsPurch = wb.Worksheet("Purchase");
                Assert.NotNull(wsPurch);

                // Row 1 is header, Rows 2..38 are data (37 records)
                Assert.Equal(38, wsPurch.LastRowUsed().RowNumber());

                // Verify Overview
                var wsOverview = wb.Worksheet("Overview");
                Assert.Equal("06ACWFS8659K1ZY (Haryana)", wsOverview.Cell("B5").GetString());
                Assert.Equal("Jan 2026", wsOverview.Cell("B6").GetString());

                // Verify sum of Taxable Value, IGST, CGST, SGST
                decimal sumTxval = 0;
                decimal sumIgst = 0;
                decimal sumCgst = 0;
                decimal sumSgst = 0;

                for (int r = 2; r <= 38; r++)
                {
                    sumTxval += (decimal)wsPurch.Cell(r, 14).GetDouble();
                    sumIgst += (decimal)wsPurch.Cell(r, 16).GetDouble();
                    sumCgst += (decimal)wsPurch.Cell(r, 17).GetDouble();
                    sumSgst += (decimal)wsPurch.Cell(r, 18).GetDouble();
                }

                Assert.Equal(24528893.38m + 165586.23m, Math.Round(sumTxval, 2));
                Assert.Equal(4213297.55m, Math.Round(sumIgst, 2));
                Assert.Equal(97013.93m + 14902.76m, Math.Round(sumCgst, 2));
                Assert.Equal(97013.93m + 14902.76m, Math.Round(sumSgst, 2));
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }

        [Fact]
        public async Task OctaGstBuilder_RealR1File_AccurateDataExtraction()
        {
            string realR1Path = @"C:\Users\ADMIN\Videos\DATA\Data for software\Data for software\Json\06ACWFS8659K1ZY_R1_012026_20260921\returns_21092026_R1_06ACWFS8659K1ZY_offline_others_0.json";
            if (!File.Exists(realR1Path)) return;

            var builder = new OctaGstBuilderService(_logger);
            string tempDir = Path.Combine(Path.GetTempPath(), $"OctaRealR1_{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);

            try
            {
                string outExcel = Path.Combine(tempDir, "Real_R1.xlsx");
                var files = new List<ScannedFileItem>
                {
                    new() { FilePath = realR1Path, FileName = Path.GetFileName(realR1Path), FileType = GstFileType.R1, Status = FileValidationStatus.Valid }
                };

                var (success, count, error) = await builder.BuildOctaWorkbookAsync(GstFileType.R1, files, outExcel);
                Assert.True(success, error);

                // 18 b2b invs + 2 cdnr + 3 hsn + 3 doc_issue = 26 records
                Assert.Equal(26, count);

                using var wb = new XLWorkbook(outExcel);
                var wsSales = wb.Worksheet("Sales");
                Assert.NotNull(wsSales);
                Assert.Equal(21, wsSales.LastRowUsed().RowNumber()); // 20 data rows + 1 header

                var wsHsn = wb.Worksheet("SalesHSN");
                Assert.NotNull(wsHsn);
                Assert.Equal(4, wsHsn.LastRowUsed().RowNumber()); // 3 data rows + 1 header

                var wsDisc = wb.Worksheet("Disclosed");
                Assert.NotNull(wsDisc);
                Assert.Equal(4, wsDisc.LastRowUsed().RowNumber()); // 3 data rows + 1 header
            }
            finally
            {
                if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true);
            }
        }
    }
}
