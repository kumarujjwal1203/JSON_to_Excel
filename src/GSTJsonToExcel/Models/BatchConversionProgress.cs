namespace GSTJsonToExcel.Models
{
    public class BatchConversionProgress
    {
        public int OverallPercentage { get; set; }
        public int CurrentFileIndex { get; set; }
        public int TotalFiles { get; set; }
        public string CurrentFileName { get; set; } = string.Empty;
        public GstFileType CurrentFileType { get; set; } = GstFileType.Unknown;
        public string Phase { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;

        public BatchConversionProgress(
            int overallPercentage,
            int currentFileIndex,
            int totalFiles,
            string currentFileName,
            GstFileType currentFileType,
            string phase,
            string message)
        {
            OverallPercentage = overallPercentage;
            CurrentFileIndex = currentFileIndex;
            TotalFiles = totalFiles;
            CurrentFileName = currentFileName;
            CurrentFileType = currentFileType;
            Phase = phase;
            Message = message;
        }
    }
}
