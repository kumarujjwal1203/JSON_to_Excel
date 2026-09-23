using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.Services.Implementations
{
    public class DuplicateDetectorService : IDuplicateDetectorService
    {
        private readonly ILoggingService _logger;

        public DuplicateDetectorService(ILoggingService logger)
        {
            _logger = logger;
        }

        public async Task<string> ComputeFileHashAsync(string filePath)
        {
            await using var stream = File.OpenRead(filePath);
            using var sha256 = SHA256.Create();
            byte[] hashBytes = await sha256.ComputeHashAsync(stream);
            return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
        }

        public async Task DetectDuplicatesAsync(List<ScannedFileItem> candidateFiles)
        {
            var seenHashes = new Dictionary<string, ScannedFileItem>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in candidateFiles)
            {
                if (item.Status != FileValidationStatus.Valid && item.Status != FileValidationStatus.UnknownType)
                {
                    continue;
                }

                try
                {
                    string hash = await ComputeFileHashAsync(item.FilePath);
                    item.ContentHash = hash;

                    if (seenHashes.TryGetValue(hash, out var original))
                    {
                        item.Status = FileValidationStatus.Duplicate;
                        item.DuplicateOf = original.FileName;
                        item.StatusReason = $"Duplicate content of '{original.FileName}' (SHA-256: {hash[..8]}...)";
                        _logger.LogInfo($"Duplicate detected: '{item.FileName}' is duplicate of '{original.FileName}'");
                    }
                    else
                    {
                        seenHashes[hash] = item;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Could not compute hash for {item.FileName}: {ex.Message}");
                }
            }
        }
    }
}
