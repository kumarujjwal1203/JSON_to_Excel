using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.Services.Implementations
{
    public class FileScannerService : IFileScannerService
    {
        private readonly IGstClassifierService _classifier;
        private readonly IDuplicateDetectorService _duplicateDetector;
        private readonly IFileService _fileService;
        private readonly ILoggingService _logger;

        public FileScannerService(
            IGstClassifierService classifier,
            IDuplicateDetectorService duplicateDetector,
            IFileService fileService,
            ILoggingService logger)
        {
            _classifier = classifier;
            _duplicateDetector = duplicateDetector;
            _fileService = fileService;
            _logger = logger;
        }

        public async Task<BatchScanSummary> ScanPathsAsync(
            IEnumerable<string> pathsOrDirectories,
            CancellationToken cancellationToken = default)
        {
            var summary = new BatchScanSummary();
            var discoveredFilePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var path in pathsOrDirectories)
            {
                if (string.IsNullOrWhiteSpace(path)) continue;

                if (Directory.Exists(path))
                {
                    CollectFilesFromDirectory(path, discoveredFilePaths);
                }
                else if (File.Exists(path))
                {
                    discoveredFilePaths.Add(path);
                }
            }

            _logger.LogInfo($"Scanner discovered {discoveredFilePaths.Count} total files from input paths.");

            foreach (var filePath in discoveredFilePaths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var fileInfo = new FileInfo(filePath);
                string ext = fileInfo.Extension.ToLowerInvariant();

                var item = new ScannedFileItem
                {
                    FilePath = filePath,
                    FileName = fileInfo.Name,
                    FileSizeBytes = fileInfo.Length,
                    Extension = ext
                };

                summary.AllDiscoveredItems.Add(item);

                // Smart filtering: Candidate JSON vs ZIP vs Ignored
                if (ext.Equals(".json", StringComparison.OrdinalIgnoreCase))
                {
                    // Validate JSON syntax & readability
                    var validation = _fileService.ValidateJsonFile(filePath);
                    if (!validation.IsValid)
                    {
                        item.Status = fileInfo.Length == 0 
                            ? FileValidationStatus.Empty 
                            : FileValidationStatus.InvalidJson;
                        item.StatusReason = validation.ErrorMessage ?? "Invalid JSON syntax.";
                    }
                    else
                    {
                        // Classify GST Return Type (R1, R3A, R2A, R2B, Unknown)
                        var (gstType, reason) = await _classifier.ClassifyFileAsync(filePath);
                        item.FileType = gstType;
                        item.StatusReason = reason;

                        if (gstType == GstFileType.Unknown)
                        {
                            item.Status = FileValidationStatus.UnknownType;
                        }
                        else
                        {
                            item.Status = FileValidationStatus.Valid;
                        }
                    }

                    summary.CandidateGstFiles.Add(item);
                }
                else if (ext.Equals(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        using var archive = ZipFile.OpenRead(filePath);
                        bool foundJsonInZip = false;

                        foreach (var entry in archive.Entries)
                        {
                            if (entry.FullName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                            {
                                foundJsonInZip = true;
                                string zipBaseName = Path.GetFileNameWithoutExtension(filePath);
                                string tempExtractDir = Path.Combine(Path.GetTempPath(), "GSTZipExtract", zipBaseName);
                                Directory.CreateDirectory(tempExtractDir);

                                string targetJsonPath = Path.Combine(tempExtractDir, $"{zipBaseName}.json");
                                entry.ExtractToFile(targetJsonPath, overwrite: true);

                                var extractedInfo = new FileInfo(targetJsonPath);
                                var extractedItem = new ScannedFileItem
                                {
                                    FilePath = targetJsonPath,
                                    FileName = $"{zipBaseName}.json",
                                    FileSizeBytes = extractedInfo.Length,
                                    Extension = ".json"
                                };

                                summary.AllDiscoveredItems.Add(extractedItem);

                                var validation = _fileService.ValidateJsonFile(targetJsonPath);
                                if (!validation.IsValid)
                                {
                                    extractedItem.Status = extractedInfo.Length == 0 ? FileValidationStatus.Empty : FileValidationStatus.InvalidJson;
                                    extractedItem.StatusReason = validation.ErrorMessage ?? "Invalid JSON syntax inside ZIP.";
                                }
                                else
                                {
                                    var (gstType, reason) = await _classifier.ClassifyFileAsync(targetJsonPath);
                                    extractedItem.FileType = gstType;
                                    extractedItem.StatusReason = $"From {fileInfo.Name}: {reason}";
                                    extractedItem.Status = gstType == GstFileType.Unknown ? FileValidationStatus.UnknownType : FileValidationStatus.Valid;
                                }

                                summary.CandidateGstFiles.Add(extractedItem);
                            }
                        }

                        if (!foundJsonInZip)
                        {
                            item.Status = FileValidationStatus.Ignored;
                            item.StatusReason = "ZIP contains no JSON return files.";
                            summary.IgnoredFiles.Add(item);
                        }
                    }
                    catch (Exception ex)
                    {
                        item.Status = FileValidationStatus.InvalidJson;
                        item.StatusReason = $"ZIP extraction error: {ex.Message}";
                        summary.CandidateGstFiles.Add(item);
                    }
                }
                else
                {
                    // Non-JSON file safely ignored without touching/deleting
                    item.Status = FileValidationStatus.Ignored;
                    item.StatusReason = $"Non-JSON file ({ext}) ignored.";
                    summary.IgnoredFiles.Add(item);
                }
            }

            // Detect Duplicates among valid candidates via SHA-256
            await _duplicateDetector.DetectDuplicatesAsync(summary.CandidateGstFiles);

            _logger.LogInfo($"Scan completed: {summary.ReadyToProcessCount} ready to convert (R1:{summary.R1Count}, R3A:{summary.R3ACount}, R2A:{summary.R2ACount}, R2B:{summary.R2BCount}), {summary.DuplicateCount} duplicates, {summary.IgnoredCount} ignored.");
            return summary;
        }

        private void CollectFilesFromDirectory(string directory, HashSet<string> collected)
        {
            try
            {
                foreach (var file in Directory.EnumerateFiles(directory, "*.*", SearchOption.AllDirectories))
                {
                    collected.Add(file);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Directory enumeration error in {directory}: {ex.Message}");
            }
        }
    }
}
