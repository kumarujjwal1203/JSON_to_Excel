using System.Collections.Generic;
using System.Linq;

namespace GSTJsonToExcel.Models
{
    public class BatchScanSummary
    {
        public List<ScannedFileItem> AllDiscoveredItems { get; } = new();
        public List<ScannedFileItem> CandidateGstFiles { get; } = new();
        public List<ScannedFileItem> IgnoredFiles { get; } = new();

        public int R1Count => CandidateGstFiles.Count(f => f.FileType == GstFileType.R1 && f.Status == FileValidationStatus.Valid);
        public int R3ACount => CandidateGstFiles.Count(f => f.FileType == GstFileType.R3A && f.Status == FileValidationStatus.Valid);
        public int R2ACount => CandidateGstFiles.Count(f => f.FileType == GstFileType.R2A && f.Status == FileValidationStatus.Valid);
        public int R2BCount => CandidateGstFiles.Count(f => f.FileType == GstFileType.R2B && f.Status == FileValidationStatus.Valid);
        public int UnknownCount => CandidateGstFiles.Count(f => f.FileType == GstFileType.Unknown && f.Status != FileValidationStatus.Duplicate && f.Status != FileValidationStatus.InvalidJson);

        public int TotalValid => CandidateGstFiles.Count(f => f.IsReadyToProcess);
        public int DuplicateCount => CandidateGstFiles.Count(f => f.Status == FileValidationStatus.Duplicate);
        public int InvalidCount => CandidateGstFiles.Count(f => f.Status == FileValidationStatus.InvalidJson || f.Status == FileValidationStatus.Empty);
        public int IgnoredCount => IgnoredFiles.Count;

        public int TotalFilesFound => AllDiscoveredItems.Count;
        public int ReadyToProcessCount => TotalValid;

        public bool HasProcessableFiles => ReadyToProcessCount > 0;
    }
}
