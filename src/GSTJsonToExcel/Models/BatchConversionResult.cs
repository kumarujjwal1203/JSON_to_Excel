using System;
using System.Collections.Generic;

namespace GSTJsonToExcel.Models
{
    public class GeneratedTypeOutput
    {
        public GstFileType FileType { get; set; }
        public string OutputFilePath { get; set; } = string.Empty;
        public string FileName => System.IO.Path.GetFileName(OutputFilePath);
        public string? SourceFileName { get; set; }
        public int FilesProcessedCount { get; set; }
        public int TotalRecordsExported { get; set; }
        public bool IntegrityPassed { get; set; }
        public long OutputFileSize { get; set; }

        public string OutputFileSizeFormatted
        {
            get
            {
                if (OutputFileSize <= 0 && System.IO.File.Exists(OutputFilePath))
                {
                    try { OutputFileSize = new System.IO.FileInfo(OutputFilePath).Length; } catch { }
                }
                double kb = (double)OutputFileSize / 1024.0;
                return kb > 1024 ? $"{kb / 1024.0:F2} MB" : $"{kb:F1} KB";
            }
        }
    }

    public class FileProcessingErrorReport
    {
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string Reason { get; set; } = string.Empty;
    }

    public class BatchConversionResult
    {
        public bool Success { get; set; }
        public string OutputDirectory { get; set; } = string.Empty;
        public List<GeneratedTypeOutput> GeneratedFiles { get; } = new();
        public List<FileProcessingErrorReport> ErrorReport { get; } = new();

        public int TotalJsonFilesProcessed { get; set; }
        public int DuplicateCount { get; set; }
        public int FailedCount { get; set; }
        public TimeSpan ElapsedTime { get; set; }

        public string SummaryMessage { get; set; } = string.Empty;
    }
}
