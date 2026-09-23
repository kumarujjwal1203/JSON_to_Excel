using System;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface ILoggingService
    {
        void LogInfo(string message);
        void LogWarning(string message);
        void LogError(string message, Exception? ex = null);
        string GetLogFilePath();
    }
}
