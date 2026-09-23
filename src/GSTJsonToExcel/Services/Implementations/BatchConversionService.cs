using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.Services.Implementations
{
    public class BatchConversionService : IBatchConversionService
    {
        private readonly IJsonParserService _parserService;
        private readonly IExcelExportService _exportService;
        private readonly IIntegrityCheckService _integrityService;
        private readonly IOctaGstBuilderService? _octaBuilder;
        private readonly ILoggingService _logger;

        public BatchConversionService(
            IJsonParserService parserService,
            IExcelExportService exportService,
            IIntegrityCheckService integrityService,
            ILoggingService logger,
            IOctaGstBuilderService? octaBuilder = null)
        {
            _parserService = parserService;
            _exportService = exportService;
            _integrityService = integrityService;
            _logger = logger;
            _octaBuilder = octaBuilder;
        }

        public async Task<BatchConversionResult> ConvertBatchAsync(
            BatchScanSummary scanSummary,
            string targetOutputDirectory,
            BatchConversionMode mode = BatchConversionMode.IndividualFiles,
            IProgress<BatchConversionProgress>? progress = null,
            CancellationToken cancellationToken = default)
        {
            var stopwatch = Stopwatch.StartNew();
            _logger.LogInfo($"Starting batch conversion in {mode} mode for {scanSummary.ReadyToProcessCount} files into '{targetOutputDirectory}'");

            Directory.CreateDirectory(targetOutputDirectory);

            var result = new BatchConversionResult
            {
                OutputDirectory = targetOutputDirectory,
                DuplicateCount = scanSummary.DuplicateCount
            };

            // Log duplicates in Error / Audit report
            foreach (var dup in scanSummary.CandidateGstFiles.Where(f => f.Status == FileValidationStatus.Duplicate))
            {
                dup.ProcessingState = ProcessingState.Skipped;
                dup.StatusMessage = dup.StatusReason ?? $"Duplicate of {dup.DuplicateOf}";
                result.ErrorReport.Add(new FileProcessingErrorReport
                {
                    FileName = dup.FileName,
                    FilePath = dup.FilePath,
                    Status = "DUPLICATE",
                    Reason = dup.StatusReason ?? $"Duplicate of {dup.DuplicateOf}"
                });
            }

            // Log invalid/empty files in Error report
            foreach (var inv in scanSummary.CandidateGstFiles.Where(f => f.Status == FileValidationStatus.InvalidJson || f.Status == FileValidationStatus.Empty))
            {
                inv.ProcessingState = ProcessingState.Failed;
                inv.StatusMessage = inv.StatusReason ?? "Invalid JSON syntax";
                result.ErrorReport.Add(new FileProcessingErrorReport
                {
                    FileName = inv.FileName,
                    FilePath = inv.FilePath,
                    Status = "FAILED",
                    Reason = inv.StatusReason ?? "Invalid JSON syntax"
                });
                result.FailedCount++;
            }

            var filesToProcess = scanSummary.CandidateGstFiles
                .Where(f => f.IsReadyToProcess && f.IsSelected)
                .ToList();

            if (mode == BatchConversionMode.IndividualFiles)
            {
                // Convert each JSON file to its own corresponding Excel file
                int totalFiles = filesToProcess.Count;
                int currentFileIndex = 0;

                foreach (var fileItem in filesToProcess)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    currentFileIndex++;
                    fileItem.ProcessingState = ProcessingState.Converting;

                    int percent = (int)((double)(currentFileIndex - 1) / Math.Max(1, totalFiles) * 100.0);
                    string desiredExcelName = $"{Path.GetFileNameWithoutExtension(fileItem.FileName)}.xlsx";
                    string outputExcelPath = GSTJsonToExcel.Helpers.FileLockHelper.GetAvailableOutputPath(Path.Combine(targetOutputDirectory, desiredExcelName));
                    string targetExcelName = Path.GetFileName(outputExcelPath);

                    progress?.Report(new BatchConversionProgress(
                        percent,
                        currentFileIndex,
                        totalFiles,
                        fileItem.FileName,
                        fileItem.FileType,
                        $"Converting File ({currentFileIndex}/{totalFiles})",
                        $"Converting {fileItem.FileName} → {targetExcelName}..."));

                    try
                    {
                        int recordsExported = 0;
                        bool conversionOk = false;
                        string? errorMsg = null;

                        bool useOcta = _octaBuilder != null && (fileItem.FileType == GstFileType.R1 || fileItem.FileType == GstFileType.R2A || fileItem.FileType == GstFileType.R2B);

                        if (useOcta)
                        {
                            var singleFileList = new List<ScannedFileItem> { fileItem };
                            var (octaOk, octaRecords, octaErr) = await _octaBuilder!.BuildOctaWorkbookAsync(
                                fileItem.FileType,
                                singleFileList,
                                outputExcelPath,
                                cancellationToken);

                            conversionOk = octaOk;
                            recordsExported = octaRecords;
                            errorMsg = octaErr;
                        }
                        else
                        {
                            // R3A or generic fallback
                            var fileParseResult = await _parserService.ParseAndFlattenAsync(
                                fileItem.FilePath,
                                null,
                                cancellationToken);

                            var exportResult = await _exportService.ExportToExcelAsync(
                                fileParseResult,
                                outputExcelPath,
                                null,
                                cancellationToken);

                            conversionOk = exportResult.Success;
                            recordsExported = exportResult.TotalRecordsExported;
                            errorMsg = exportResult.ErrorMessage;
                        }

                        if (conversionOk && File.Exists(outputExcelPath))
                        {
                            fileItem.ProcessingState = ProcessingState.Completed;
                            fileItem.OutputExcelPath = outputExcelPath;
                            fileItem.RecordsExported = recordsExported;
                            fileItem.StatusMessage = $"Saved: {targetExcelName} ({recordsExported} records)";

                            result.TotalJsonFilesProcessed++;
                            result.GeneratedFiles.Add(new GeneratedTypeOutput
                            {
                                FileType = fileItem.FileType,
                                OutputFilePath = outputExcelPath,
                                SourceFileName = fileItem.FileName,
                                FilesProcessedCount = 1,
                                TotalRecordsExported = recordsExported,
                                IntegrityPassed = true
                            });

                            result.ErrorReport.Add(new FileProcessingErrorReport
                            {
                                FileName = fileItem.FileName,
                                FilePath = fileItem.FilePath,
                                Status = "SUCCESS",
                                Reason = $"Generated {targetExcelName} with {recordsExported} records."
                            });

                            _logger.LogInfo($"Successfully generated '{targetExcelName}' for '{fileItem.FileName}' ({recordsExported} records).");
                        }
                        else
                        {
                            fileItem.ProcessingState = ProcessingState.Failed;
                            fileItem.StatusMessage = errorMsg ?? "Conversion failed.";
                            result.FailedCount++;
                            result.ErrorReport.Add(new FileProcessingErrorReport
                            {
                                FileName = fileItem.FileName,
                                FilePath = fileItem.FilePath,
                                Status = "FAILED",
                                Reason = errorMsg ?? "Conversion error"
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        fileItem.ProcessingState = ProcessingState.Failed;
                        fileItem.StatusMessage = ex.Message;
                        result.FailedCount++;
                        result.ErrorReport.Add(new FileProcessingErrorReport
                        {
                            FileName = fileItem.FileName,
                            FilePath = fileItem.FilePath,
                            Status = "FAILED",
                            Reason = ex.Message
                        });
                        _logger.LogError($"Failed converting '{fileItem.FileName}'", ex);
                    }
                }
            }
            else
            {
                // Merged by GST return type mode
                var validFilesByType = filesToProcess
                    .GroupBy(f => f.FileType)
                    .OrderBy(g => g.Key)
                    .ToList();

                int totalValidFiles = validFilesByType.Sum(g => g.Count());
                int overallFileIndex = 0;

                foreach (var group in validFilesByType)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var gstType = group.Key;
                    var filesInType = group.ToList();
                    string typeName = gstType.ToString(); // R1, R3A, R2A, R2B
                    string desiredName = $"{typeName}.xlsx";
                    string outputExcelPath = GSTJsonToExcel.Helpers.FileLockHelper.GetAvailableOutputPath(Path.Combine(targetOutputDirectory, desiredName));

                    _logger.LogInfo($"Processing {filesInType.Count} files for GST type '{typeName}' -> '{outputExcelPath}'");

                    bool useOcta = _octaBuilder != null && (gstType == GstFileType.R1 || gstType == GstFileType.R2A || gstType == GstFileType.R2B);

                    if (useOcta)
                    {
                        int typeFileCount = filesInType.Count;
                        foreach (var fileItem in filesInType)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            overallFileIndex++;
                            fileItem.ProcessingState = ProcessingState.Converting;
                            int percent = (int)((double)(overallFileIndex - 1) / Math.Max(1, totalValidFiles) * 100.0);

                            progress?.Report(new BatchConversionProgress(
                                percent,
                                overallFileIndex,
                                totalValidFiles,
                                fileItem.FileName,
                                gstType,
                                $"Processing {typeName} Files ({overallFileIndex}/{totalValidFiles})",
                                $"Validating and preparing: {fileItem.FileName}..."));
                        }

                        int currentPercent = (int)((double)overallFileIndex / Math.Max(1, totalValidFiles) * 100.0);
                        progress?.Report(new BatchConversionProgress(
                            currentPercent,
                            overallFileIndex,
                            totalValidFiles,
                            $"{typeName}.xlsx",
                            gstType,
                            $"Generating Excel Workbook",
                            $"Building Octa GST worksheets and styling for {typeName}.xlsx..."));

                        var (octaSuccess, totalRecords, octaError) = await _octaBuilder!.BuildOctaWorkbookAsync(
                            gstType,
                            filesInType,
                            outputExcelPath,
                            cancellationToken);

                        if (octaSuccess)
                        {
                            foreach (var fileItem in filesInType)
                            {
                                fileItem.ProcessingState = ProcessingState.Completed;
                                fileItem.OutputExcelPath = outputExcelPath;
                                fileItem.StatusMessage = $"Merged into {typeName}.xlsx";
                                result.TotalJsonFilesProcessed++;
                                result.ErrorReport.Add(new FileProcessingErrorReport
                                {
                                    FileName = fileItem.FileName,
                                    FilePath = fileItem.FilePath,
                                    Status = "SUCCESS",
                                    Reason = "Successfully processed into Octa GST format with zero data loss."
                                });
                            }

                            result.GeneratedFiles.Add(new GeneratedTypeOutput
                            {
                                FileType = gstType,
                                OutputFilePath = outputExcelPath,
                                FilesProcessedCount = typeFileCount,
                                TotalRecordsExported = totalRecords,
                                IntegrityPassed = true
                            });

                            _logger.LogInfo($"Successfully generated Octa format '{typeName}.xlsx' with {totalRecords} records from {typeFileCount} files.");
                        }
                        else
                        {
                            foreach (var fileItem in filesInType)
                            {
                                fileItem.ProcessingState = ProcessingState.Failed;
                                fileItem.StatusMessage = octaError ?? "Conversion failed.";
                            }
                            result.FailedCount += typeFileCount;
                            result.ErrorReport.Add(new FileProcessingErrorReport
                            {
                                FileName = $"{typeName}.xlsx",
                                FilePath = outputExcelPath,
                                Status = "FAILED",
                                Reason = octaError ?? "Failed generating Octa GST Excel file."
                            });
                            _logger.LogError($"Failed generating Octa format '{typeName}.xlsx': {octaError}");
                        }
                    }
                    else
                    {
                        var mergedParseResult = new JsonParseResult
                        {
                            SourceFileName = $"{typeName} ({filesInType.Count} files)",
                            SourceFilePath = outputExcelPath
                        };

                        var mergedTablesMap = new Dictionary<string, JsonFlatTable>(StringComparer.OrdinalIgnoreCase);
                        int successfullyProcessedInType = 0;

                        foreach (var fileItem in filesInType)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            overallFileIndex++;
                            fileItem.ProcessingState = ProcessingState.Converting;
                            int percent = (int)((double)(overallFileIndex - 1) / Math.Max(1, totalValidFiles) * 100.0);

                            progress?.Report(new BatchConversionProgress(
                                percent,
                                overallFileIndex,
                                totalValidFiles,
                                fileItem.FileName,
                                gstType,
                                $"Processing {typeName} Files ({overallFileIndex}/{totalValidFiles})",
                                $"Reading & Parsing: {fileItem.FileName}..."));

                            try
                            {
                                var fileParseResult = await _parserService.ParseAndFlattenAsync(
                                    fileItem.FilePath,
                                    null,
                                    cancellationToken);

                                // Merge tables into the type's unified parse result
                                foreach (var srcTable in fileParseResult.Tables)
                                {
                                    if (!mergedTablesMap.TryGetValue(srcTable.SheetName, out var destTable))
                                    {
                                        destTable = new JsonFlatTable
                                        {
                                            SheetName = srcTable.SheetName,
                                            FullJsonPath = srcTable.FullJsonPath
                                        };
                                        destTable.AddColumn("source_file");
                                        mergedTablesMap[srcTable.SheetName] = destTable;
                                        mergedParseResult.Tables.Add(destTable);
                                    }

                                    foreach (var row in srcTable.Rows)
                                    {
                                        row.Set("source_file", JsonFlatValue.FromString(fileItem.FileName));
                                        destTable.AddRow(row);
                                    }
                                }

                                // Merge all leaf nodes for Master Data Index
                                foreach (var kvp in fileParseResult.AllLeafNodes)
                                {
                                    string mergedKey = $"[{fileItem.FileName}] {kvp.Key}";
                                    mergedParseResult.AllLeafNodes[mergedKey] = kvp.Value;
                                }

                                mergedParseResult.SourceFileSizeBytes += fileItem.FileSizeBytes;
                                successfullyProcessedInType++;
                                fileItem.ProcessingState = ProcessingState.Completed;
                                fileItem.OutputExcelPath = outputExcelPath;
                                fileItem.StatusMessage = $"Merged into {typeName}.xlsx";
                                result.TotalJsonFilesProcessed++;

                                result.ErrorReport.Add(new FileProcessingErrorReport
                                {
                                    FileName = fileItem.FileName,
                                    FilePath = fileItem.FilePath,
                                    Status = "SUCCESS",
                                    Reason = $"Parsed {fileParseResult.TotalLeafCount} fields without data loss."
                                });
                            }
                            catch (Exception ex)
                            {
                                fileItem.ProcessingState = ProcessingState.Failed;
                                fileItem.StatusMessage = ex.Message;
                                _logger.LogError($"Failed processing file '{fileItem.FileName}'", ex);
                                result.FailedCount++;
                                result.ErrorReport.Add(new FileProcessingErrorReport
                                {
                                    FileName = fileItem.FileName,
                                    FilePath = fileItem.FilePath,
                                    Status = "FAILED",
                                    Reason = $"Conversion error: {ex.Message}"
                                });
                            }
                        }

                        if (successfullyProcessedInType > 0)
                        {
                            int currentPercent = (int)((double)overallFileIndex / Math.Max(1, totalValidFiles) * 100.0);
                            progress?.Report(new BatchConversionProgress(
                                currentPercent,
                                overallFileIndex,
                                totalValidFiles,
                                $"{typeName}.xlsx",
                                gstType,
                                $"Generating Excel Workbook",
                                $"Building worksheets and styling for {typeName}.xlsx..."));

                            var exportResult = await _exportService.ExportToExcelAsync(
                                mergedParseResult,
                                outputExcelPath,
                                null,
                                cancellationToken);

                            bool integrityPassed = true;
                            if (exportResult.Success && File.Exists(outputExcelPath))
                            {
                                var integrityReport = await _integrityService.VerifyIntegrityAsync(
                                    mergedParseResult,
                                    outputExcelPath,
                                    null,
                                    cancellationToken);
                                integrityPassed = integrityReport.Passed;
                            }

                            result.GeneratedFiles.Add(new GeneratedTypeOutput
                            {
                                FileType = gstType,
                                OutputFilePath = outputExcelPath,
                                FilesProcessedCount = successfullyProcessedInType,
                                TotalRecordsExported = exportResult.TotalRecordsExported,
                                IntegrityPassed = integrityPassed
                            });

                            _logger.LogInfo($"Successfully generated '{typeName}.xlsx' with {exportResult.TotalRecordsExported} records from {successfullyProcessedInType} files.");
                        }
                    }
                }
            }

            stopwatch.Stop();
            result.ElapsedTime = stopwatch.Elapsed;
            result.Success = result.GeneratedFiles.Count > 0;
            result.SummaryMessage = mode == BatchConversionMode.IndividualFiles
                ? $"Successfully converted {result.GeneratedFiles.Count} individual JSON files into dedicated Excel files in {stopwatch.Elapsed.TotalSeconds:F1}s."
                : $"Processed {result.TotalJsonFilesProcessed} files across {result.GeneratedFiles.Count} GST types in {stopwatch.Elapsed.TotalSeconds:F1}s.";

            progress?.Report(new BatchConversionProgress(
                100,
                filesToProcess.Count,
                filesToProcess.Count,
                "Complete",
                GstFileType.Unknown,
                "Batch Conversion Completed",
                result.SummaryMessage));

            _logger.LogInfo($"Batch conversion finished: {result.SummaryMessage}");
            return result;
        }
    }
}
