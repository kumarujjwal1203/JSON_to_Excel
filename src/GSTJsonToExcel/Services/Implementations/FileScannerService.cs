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
            return await Task.Run(async () =>
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

                        var (periodLabel, sortKey) = ExtractReturnPeriod(filePath, item.FileName);
                        item.ReturnPeriod = periodLabel;
                        item.PeriodSortKey = sortKey;
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

                                    var (periodLabel, sortKey) = ExtractReturnPeriod(targetJsonPath, extractedItem.FileName);
                                    extractedItem.ReturnPeriod = periodLabel;
                                    extractedItem.PeriodSortKey = sortKey;
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
            }, cancellationToken);
        }

        private static readonly string[] MonthShortNames =
        {
            "", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        };

        private static (string Label, int SortKey) ExtractReturnPeriod(string filePath, string fileName)
        {
            try
            {
                using var stream = File.OpenRead(filePath);
                using var doc = System.Text.Json.JsonDocument.Parse(stream);
                var root = doc.RootElement;
                if (root.ValueKind == System.Text.Json.JsonValueKind.Object)
                {
                    string? rawFp = TryGetPeriodProperty(root);
                    if (string.IsNullOrWhiteSpace(rawFp) &&
                        root.TryGetProperty("data", out var dataEl) &&
                        dataEl.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        rawFp = TryGetPeriodProperty(dataEl);
                    }

                    if (TryParseMmYyyy(rawFp, out int m, out int y))
                    {
                        return ($"{MonthShortNames[m]} {y}", y * 100 + m);
                    }
                }
            }
            catch
            {
                // Fallback to filename
            }

            var match = System.Text.RegularExpressions.Regex.Match(fileName ?? string.Empty, @"(?<!\d)(0[1-9]|1[0-2])(20\d{2})(?!\d)");
            if (match.Success &&
                int.TryParse(match.Groups[1].Value, out int fm) &&
                int.TryParse(match.Groups[2].Value, out int fy))
            {
                return ($"{MonthShortNames[fm]} {fy}", fy * 100 + fm);
            }

            return ("Other / General", 999999);
        }

        private static string? TryGetPeriodProperty(System.Text.Json.JsonElement obj)
        {
            string[] keys = { "fp", "ret_period", "rtnprd", "taxperiod" };
            foreach (var prop in obj.EnumerateObject())
            {
                foreach (var k in keys)
                {
                    if (string.Equals(prop.Name, k, StringComparison.OrdinalIgnoreCase) &&
                        prop.Value.ValueKind == System.Text.Json.JsonValueKind.String)
                    {
                        return prop.Value.GetString();
                    }
                }
            }
            return null;
        }

        private static bool TryParseMmYyyy(string? raw, out int month, out int year)
        {
            month = 0;
            year = 0;
            if (string.IsNullOrWhiteSpace(raw)) return false;
            string digits = raw.Trim();
            if (digits.Length == 6 &&
                int.TryParse(digits.Substring(0, 2), out int m) &&
                int.TryParse(digits.Substring(2, 4), out int y) &&
                m >= 1 && m <= 12 && y >= 2017 && y <= 2099)
            {
                month = m;
                year = y;
                return true;
            }
            return false;
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
