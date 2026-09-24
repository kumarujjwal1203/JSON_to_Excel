using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Features.AnnualReport.Models;
using GSTJsonToExcel.Helpers;

namespace GSTJsonToExcel.Features.AnnualReport.Services
{
    public class AnnualReportScannerService
    {
        public async Task<List<AnnualScannedFile>> ScanPathsAsync(
            IEnumerable<string> inputPaths,
            CancellationToken cancellationToken = default)
        {
            return await Task.Run(() =>
            {
                var candidateFiles = new List<string>();
                foreach (var path in inputPaths)
                {
                    if (string.IsNullOrWhiteSpace(path)) continue;
                    if (Directory.Exists(path))
                    {
                        try
                        {
                            candidateFiles.AddRange(Directory.EnumerateFiles(path, "*.*", SearchOption.AllDirectories)
                                .Where(f => f.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ||
                                            f.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)));
                        }
                        catch
                        {
                            // Ignore inaccessible subdirectories
                        }
                    }
                    else if (File.Exists(path))
                    {
                        if (path.EndsWith(".json", StringComparison.OrdinalIgnoreCase) ||
                            path.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                        {
                            candidateFiles.Add(path);
                        }
                    }
                }

                var results = new List<AnnualScannedFile>();
                foreach (var filePath in candidateFiles.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (filePath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        ScanZipFile(filePath, results);
                    }
                    else
                    {
                        ScanSingleJsonFile(filePath, results);
                    }
                }

                MarkDuplicates(results);
                return results;
            }, cancellationToken);
        }

        private void ScanZipFile(string zipPath, List<AnnualScannedFile> results)
        {
            try
            {
                var fileInfo = new FileInfo(zipPath);
                using var archive = ZipFile.OpenRead(zipPath);
                var jsonEntries = archive.Entries
                    .Where(e => e.FullName.EndsWith(".json", StringComparison.OrdinalIgnoreCase) && e.Length > 0)
                    .ToList();

                if (jsonEntries.Count == 0)
                {
                    results.Add(new AnnualScannedFile
                    {
                        SourcePath = zipPath,
                        DisplayFileName = Path.GetFileName(zipPath),
                        IsCorrupt = true,
                        StatusReason = "ZIP contains no JSON files",
                        LastModified = fileInfo.LastWriteTimeUtc
                    });
                    return;
                }

                foreach (var entry in jsonEntries)
                {
                    try
                    {
                        using var stream = entry.Open();
                        var item = InspectJsonStream(stream, zipPath, entry.FullName, fileInfo.LastWriteTimeUtc);
                        results.Add(item);
                    }
                    catch (Exception ex)
                    {
                        results.Add(new AnnualScannedFile
                        {
                            SourcePath = zipPath,
                            ZipEntryPath = entry.FullName,
                            DisplayFileName = $"{Path.GetFileName(zipPath)}/{entry.Name}",
                            IsCorrupt = true,
                            StatusReason = $"Corrupt JSON inside ZIP: {ex.Message}",
                            LastModified = fileInfo.LastWriteTimeUtc
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                results.Add(new AnnualScannedFile
                {
                    SourcePath = zipPath,
                    DisplayFileName = Path.GetFileName(zipPath),
                    IsCorrupt = true,
                    StatusReason = $"Invalid ZIP archive: {ex.Message}"
                });
            }
        }

        private void ScanSingleJsonFile(string jsonPath, List<AnnualScannedFile> results)
        {
            try
            {
                var fileInfo = new FileInfo(jsonPath);
                using var stream = new FileStream(jsonPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var item = InspectJsonStream(stream, jsonPath, null, fileInfo.LastWriteTimeUtc);
                results.Add(item);
            }
            catch (Exception ex)
            {
                results.Add(new AnnualScannedFile
                {
                    SourcePath = jsonPath,
                    DisplayFileName = Path.GetFileName(jsonPath),
                    IsCorrupt = true,
                    StatusReason = $"Corrupt JSON: {ex.Message}"
                });
            }
        }

        private AnnualScannedFile InspectJsonStream(
            Stream stream,
            string sourcePath,
            string? zipEntryPath,
            DateTime lastModifiedUtc)
        {
            string displayFileName = zipEntryPath == null
                ? Path.GetFileName(sourcePath)
                : $"{Path.GetFileName(sourcePath)}!{Path.GetFileName(zipEntryPath)}";

            var item = new AnnualScannedFile
            {
                SourcePath = sourcePath,
                ZipEntryPath = zipEntryPath,
                DisplayFileName = displayFileName,
                LastModified = lastModifiedUtc
            };

            using var doc = JsonDocument.Parse(stream);
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
            {
                item.IsCorrupt = true;
                item.StatusReason = "JSON root is not an object";
                return item;
            }

            // Detect if wrapped in "data" (standard for GSTR-2B or some offline payloads)
            JsonElement payload = root;
            bool hasDataObject = root.TryGetProperty("data", out var dataProp) && dataProp.ValueKind == JsonValueKind.Object;
            if (hasDataObject)
            {
                payload = dataProp;
            }

            // 1. Extract GSTIN (with fallback to filename/path regex)
            string gstin = GetString(root, "gstin");
            if (string.IsNullOrWhiteSpace(gstin) && hasDataObject)
            {
                gstin = GetString(payload, "gstin");
            }
            if (string.IsNullOrWhiteSpace(gstin))
            {
                var mGstin = System.Text.RegularExpressions.Regex.Match(
                    $"{displayFileName} {sourcePath}",
                    @"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (mGstin.Success) gstin = mGstin.Value;
            }
            item.Gstin = gstin.Trim().ToUpperInvariant();
            item.StateName = GstStateHelper.GetStateName(item.Gstin);

            // 2. Extract Return Period (MMYYYY or filename fallback)
            string period = GetString(root, "ret_period");
            if (string.IsNullOrWhiteSpace(period)) period = GetString(root, "fp");
            if (string.IsNullOrWhiteSpace(period)) period = GetString(root, "rtnprd");
            if (string.IsNullOrWhiteSpace(period)) period = GetString(root, "taxperiod");
            if (string.IsNullOrWhiteSpace(period) && hasDataObject)
            {
                period = GetString(payload, "rtnprd");
                if (string.IsNullOrWhiteSpace(period)) period = GetString(payload, "fp");
                if (string.IsNullOrWhiteSpace(period)) period = GetString(payload, "ret_period");
                if (string.IsNullOrWhiteSpace(period)) period = GetString(payload, "taxperiod");
            }
            item.ReturnPeriod = NormalizePeriodToMmYyyy(period, displayFileName);

            // 3. Parse Month & FY
            if (item.ReturnPeriod.Length == 6 &&
                int.TryParse(item.ReturnPeriod[..2], out int mm) &&
                int.TryParse(item.ReturnPeriod[2..], out int yyyy) &&
                mm >= 1 && mm <= 12 && yyyy >= 2017 && yyyy <= 2100)
            {
                item.MonthNumber = mm;
                item.YearNumber = yyyy;
                int fyStartYear = mm >= 4 ? yyyy : yyyy - 1;
                int fyEndShort = (fyStartYear + 1) % 100;
                item.FinancialYear = $"{fyStartYear}-{fyEndShort:D2}";
                item.FyMonthIndex = mm >= 4 ? mm - 4 : mm + 8; // Apr=0 .. Mar=11
            }

            // 4. Detect Return Type by content keys + filename hint
            item.ReturnType = DetectReturnType(root, payload, hasDataObject, sourcePath, zipEntryPath);

            if (item.ReturnType == AnnualReturnType.Unknown || string.IsNullOrWhiteSpace(item.Gstin) || item.FyMonthIndex < 0)
            {
                item.IsCorrupt = true;
                item.StatusReason = "Unrecognized GST return format or missing GSTIN/period";
            }

            return item;
        }

        private static string NormalizePeriodToMmYyyy(string rawPeriod, string fileName)
        {
            string p = (rawPeriod ?? string.Empty).Trim().Replace("-", "").Replace("/", "").Replace(" ", "");
            if (p.Length == 6 && int.TryParse(p[..2], out int m1) && int.TryParse(p[2..], out int y1) && m1 >= 1 && m1 <= 12 && y1 >= 2017 && y1 <= 2100)
            {
                return $"{m1:D2}{y1}";
            }
            if (p.Length == 6 && int.TryParse(p[..4], out int y2) && int.TryParse(p[4..], out int m2) && m2 >= 1 && m2 <= 12 && y2 >= 2017 && y2 <= 2100)
            {
                return $"{m2:D2}{y2}";
            }

            // Fallback 1: 6-digit MMYYYY in filename (e.g. _042025_)
            var m6 = System.Text.RegularExpressions.Regex.Match(fileName ?? string.Empty, @"(?<!\d)(0[1-9]|1[0-2])(20\d{2})(?!\d)");
            if (m6.Success)
            {
                return $"{m6.Groups[1].Value}{m6.Groups[2].Value}";
            }

            // Fallback 2: 8-digit DDMMYYYY in GST Offline Tool filename (e.g. returns_18092026_R1_...)
            var m8 = System.Text.RegularExpressions.Regex.Match(fileName ?? string.Empty, @"(?<!\d)(?:0[1-9]|[12]\d|3[01])(0[1-9]|1[0-2])(20\d{2})(?!\d)");
            if (m8.Success)
            {
                return $"{m8.Groups[1].Value}{m8.Groups[2].Value}";
            }

            return p;
        }

        private static AnnualReturnType DetectReturnType(
            JsonElement root,
            JsonElement payload,
            bool hasDataObject,
            string sourcePath,
            string? zipEntryPath)
        {
            // GSTR-2B check: has docdata or itcsumm or data.rtnprd
            if (payload.TryGetProperty("docdata", out _) ||
                payload.TryGetProperty("itcsumm", out _) ||
                payload.TryGetProperty("cpsumm", out _) ||
                root.TryGetProperty("docdata", out _))
            {
                return AnnualReturnType.GSTR2B;
            }

            // GSTR-3B check: has sup_details, itc_elg, inward_sup, intr_ltfee, taxpayble, tx_pmt
            if (payload.TryGetProperty("sup_details", out _) ||
                payload.TryGetProperty("itc_elg", out _) ||
                payload.TryGetProperty("inward_sup", out _) ||
                payload.TryGetProperty("taxpayble", out _) ||
                payload.TryGetProperty("tx_pmt", out _) ||
                payload.TryGetProperty("intr_ltfee", out _))
            {
                return AnnualReturnType.GSTR3B;
            }

            // GSTR-1 exclusive keys
            if (payload.TryGetProperty("b2cs", out _) ||
                payload.TryGetProperty("b2cl", out _) ||
                payload.TryGetProperty("exp", out _) ||
                payload.TryGetProperty("cdnr", out _) ||
                payload.TryGetProperty("cdnur", out _) ||
                payload.TryGetProperty("hsn", out _) ||
                payload.TryGetProperty("doc_issue", out _) ||
                payload.TryGetProperty("nil", out _) ||
                payload.TryGetProperty("txpd", out _) ||
                payload.TryGetProperty("at", out _))
            {
                return AnnualReturnType.GSTR1;
            }

            // GSTR-2A exclusive keys
            if (payload.TryGetProperty("cdn", out _) ||
                payload.TryGetProperty("tdsa", out _) ||
                payload.TryGetProperty("tds", out _) ||
                payload.TryGetProperty("tcs", out _) ||
                payload.TryGetProperty("impg", out _) ||
                payload.TryGetProperty("impgsez", out _))
            {
                return AnnualReturnType.GSTR2A;
            }

            // If only "b2b" or "b2ba" is present, inspect filename or inner b2b structure
            string combinedName = $"{Path.GetFileName(sourcePath)} {zipEntryPath}".ToUpperInvariant();
            if (combinedName.Contains("R2A") || combinedName.Contains("GSTR2A") || combinedName.Contains("2A"))
            {
                return AnnualReturnType.GSTR2A;
            }
            if (combinedName.Contains("R1") || combinedName.Contains("GSTR1") || combinedName.Contains("returns_"))
            {
                return AnnualReturnType.GSTR1;
            }
            if (combinedName.Contains("R3B") || combinedName.Contains("GSTR3B"))
            {
                return AnnualReturnType.GSTR3B;
            }
            if (combinedName.Contains("R2B") || combinedName.Contains("GSTR2B"))
            {
                return AnnualReturnType.GSTR2B;
            }

            if (payload.TryGetProperty("b2b", out _))
            {
                return AnnualReturnType.GSTR1;
            }

            return AnnualReturnType.Unknown;
        }

        private static void MarkDuplicates(List<AnnualScannedFile> files)
        {
            var validGroups = files
                .Where(f => !f.IsCorrupt && f.ReturnType != AnnualReturnType.Unknown)
                .GroupBy(f => $"{f.Gstin}|{f.ReturnType}|{f.ReturnPeriod}", StringComparer.OrdinalIgnoreCase);

            foreach (var group in validGroups)
            {
                var ordered = group
                    .OrderByDescending(f => f.LastModified)
                    .ThenByDescending(f => f.ZipEntryPath != null) // Prefer official portal zip if timestamps match
                    .ToList();

                for (int i = 1; i < ordered.Count; i++)
                {
                    ordered[i].IsDuplicate = true;
                    ordered[i].StatusReason = "Duplicate (older timestamp skipped)";
                }
            }
        }

        public static Stream OpenFileStream(AnnualScannedFile file)
        {
            if (string.IsNullOrEmpty(file.ZipEntryPath))
            {
                return new FileStream(file.SourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }

            var memoryStream = new MemoryStream();
            using (var archive = ZipFile.OpenRead(file.SourcePath))
            {
                var entry = archive.GetEntry(file.ZipEntryPath)
                            ?? archive.Entries.FirstOrDefault(e => string.Equals(e.FullName, file.ZipEntryPath, StringComparison.OrdinalIgnoreCase));
                if (entry != null)
                {
                    using var entryStream = entry.Open();
                    entryStream.CopyTo(memoryStream);
                }
            }
            memoryStream.Position = 0;
            return memoryStream;
        }

        private static string GetString(JsonElement el, string propName)
        {
            if (el.ValueKind == JsonValueKind.Object &&
                el.TryGetProperty(propName, out var val) &&
                val.ValueKind == JsonValueKind.String)
            {
                return val.GetString() ?? string.Empty;
            }
            return string.Empty;
        }
    }
}
