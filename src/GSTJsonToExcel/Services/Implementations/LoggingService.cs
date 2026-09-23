using System;
using System.IO;
using GSTJsonToExcel.Services.Interfaces;

namespace GSTJsonToExcel.Services.Implementations
{
    public class LoggingService : ILoggingService
    {
        private readonly string _logFilePath;
        private readonly object _lock = new();

        public LoggingService()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logDir = Path.Combine(appData, "GSTJsonToExcel", "logs");
            Directory.CreateDirectory(logDir);
            _logFilePath = Path.Combine(logDir, "app.log");
        }

        public string GetLogFilePath() => _logFilePath;

        public void LogInfo(string message)
        {
            WriteEntry("INFO", message);
        }

        public void LogWarning(string message)
        {
            WriteEntry("WARN", message);
        }

        public void LogError(string message, Exception? ex = null)
        {
            string fullMessage = ex == null ? message : $"{message} | Exception: {ex.GetType().Name}: {ex.Message}";
            WriteEntry("ERROR", fullMessage);
        }

        private void WriteEntry(string level, string message)
        {
            try
            {
                lock (_lock)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    File.AppendAllText(_logFilePath, $"{timestamp} [{level}] {message}{Environment.NewLine}");
                }
            }
            catch
            {
                // Silently ignore logging failures to never crash the main application
            }
        }
    }
}
