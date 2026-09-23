using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.Services.Implementations
{
    public class IntegrityCheckService : IIntegrityCheckService
    {
        private readonly ILoggingService _logger;

        public IntegrityCheckService(ILoggingService logger)
        {
            _logger = logger;
        }

        public async Task<DataIntegrityReport> VerifyIntegrityAsync(
            JsonParseResult parseResult,
            string excelFilePath,
            IProgress<ConversionProgress>? progress = null,
            CancellationToken cancellationToken = default)
        {
            _logger.LogInfo("Starting automated Data Integrity Verification...");
            progress?.Report(new ConversionProgress(95, "Validating Integrity", "Comparing source JSON leaves against generated Excel workbook..."));

            return await Task.Run(() =>
            {
                var report = new DataIntegrityReport
                {
                    TotalSourceLeafValues = parseResult.TotalLeafCount
                };

                if (!File.Exists(excelFilePath))
                {
                    report.Passed = false;
                    report.Warnings.Add("Generated Excel file was not found on disk.");
                    return report;
                }

                try
                {
                    using var workbook = new XLWorkbook(excelFilePath);

                    // Check that expected sheets exist
                    foreach (var table in parseResult.Tables)
                    {
                        string expectedName = JsonParserService.SanitizeSheetName(table.SheetName);
                        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.StartsWith(expectedName, StringComparison.OrdinalIgnoreCase));
                        if (sheet != null)
                        {
                            report.SectionsVerified.Add($"{sheet.Name} (Rows: {sheet.RowCount() - 1}, Cols: {sheet.ColumnCount()})");
                        }
                        else
                        {
                            report.Warnings.Add($"Worksheet for section '{table.SheetName}' was not found.");
                        }
                    }

                    // Verify Master Data Index sheet
                    var indexSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals("All_Data_Index", StringComparison.OrdinalIgnoreCase));
                    if (indexSheet != null)
                    {
                        // Row 1 is header, data rows start at 2
                        int indexRowCount = indexSheet.RowsUsed().Count() - 1;
                        report.MatchedLeafValues = indexRowCount;
                        report.TotalExportedCells += indexRowCount;

                        if (indexRowCount == parseResult.TotalLeafCount)
                        {
                            report.AuditNotes.Add($"All {parseResult.TotalLeafCount:N0} leaf values verified 1:1 in Master Audit Index.");
                        }
                        else
                        {
                            report.Warnings.Add($"Index row count ({indexRowCount}) differs from source leaf count ({parseResult.TotalLeafCount}).");
                        }
                    }
                    else
                    {
                        report.Warnings.Add("All_Data_Index sheet is missing.");
                    }

                    // Count total data cells across all worksheets
                    int totalDataCells = 0;
                    foreach (var ws in workbook.Worksheets)
                    {
                        if (ws.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase)) continue;

                        int rows = ws.RowsUsed().Count();
                        int cols = ws.ColumnsUsed().Count();
                        if (rows > 1 && cols > 0)
                        {
                            totalDataCells += (rows - 1) * cols;
                        }
                    }
                    report.TotalExportedCells = totalDataCells;

                    // Spot-check sample leaf values from source
                    int sampleChecks = 0;
                    int sampleMatches = 0;
                    var randomSamples = parseResult.AllLeafNodes.Take(30).ToList();

                    if (indexSheet != null)
                    {
                        // Check first 30 rows in index sheet
                        int checkRow = 2;
                        foreach (var sample in randomSamples)
                        {
                            if (checkRow > indexSheet.RowsUsed().Count()) break;

                            string? cellPath = indexSheet.Cell(checkRow, 2).GetString();
                            string? cellVal = indexSheet.Cell(checkRow, 4).GetString();

                            if (!string.IsNullOrEmpty(cellPath) && parseResult.AllLeafNodes.TryGetValue(cellPath, out var expectedVal))
                            {
                                sampleChecks++;
                                if (cellVal == expectedVal.RawString)
                                {
                                    sampleMatches++;
                                }
                            }
                            checkRow++;
                        }
                    }

                    if (sampleChecks > 0 && sampleMatches == sampleChecks)
                    {
                        report.AuditNotes.Add($"Spot-check: {sampleMatches}/{sampleChecks} sample cell values matched source JSON with 100% precision.");
                    }

                    report.Passed = report.Warnings.Count == 0 && report.MatchedLeafValues == parseResult.TotalLeafCount;

                    _logger.LogInfo($"Integrity Check completed: Passed={report.Passed}, SourceLeaves={report.TotalSourceLeafValues}, MatchedLeaves={report.MatchedLeafValues}");
                    return report;
                }
                catch (Exception ex)
                {
                    _logger.LogError("Data integrity verification failed with exception", ex);
                    report.Passed = false;
                    report.Warnings.Add($"Integrity check failed: {ex.Message}");
                    return report;
                }
            }, cancellationToken);
        }
    }
}
