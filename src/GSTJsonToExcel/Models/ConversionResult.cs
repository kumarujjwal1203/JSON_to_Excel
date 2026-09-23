using System;

namespace GSTJsonToExcel.Models
{
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public string OutputFilePath { get; set; } = string.Empty;
        public long FileSizeBytes { get; set; }
        public int TotalSheetsCreated { get; set; }
        public int TotalRecordsExported { get; set; }
        public int TotalFieldsExported { get; set; }
        public DataIntegrityReport IntegrityReport { get; set; } = new();
        public TimeSpan ElapsedTime { get; set; }

        public static ConversionResult Failure(string error)
        {
            return new ConversionResult
            {
                Success = false,
                ErrorMessage = error
            };
        }
    }
}
