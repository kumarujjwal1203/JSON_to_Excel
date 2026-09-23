namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IFileService
    {
        string? SelectJsonFile();
        string[]? SelectJsonFiles();
        string? SelectFolder();
        string? SelectSaveLocation(string defaultFileName);
        void OpenFile(string filePath);
        void OpenFolder(string filePath);
        (bool IsValid, string? ErrorMessage, long FileSizeBytes) ValidateJsonFile(string filePath);
    }
}
