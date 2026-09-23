using System;
using System.IO;

namespace GSTJsonToExcel.Helpers
{
    public static class FileLockHelper
    {
        /// <summary>
        /// Ensures an output path is writable. If the target file is locked by Excel or another process,
        /// generates an alternative available file path (_v1, _v2, etc.) to prevent save crashes.
        /// </summary>
        public static string GetAvailableOutputPath(string desiredPath)
        {
            if (string.IsNullOrWhiteSpace(desiredPath)) return desiredPath;

            if (!File.Exists(desiredPath)) return desiredPath;

            // Check if file is writable (not locked by Excel)
            try
            {
                using var stream = File.Open(desiredPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return desiredPath; // Not locked, can be overwritten safely
            }
            catch (IOException)
            {
                // Locked by another process (e.g. Microsoft Excel). Find safe alternate path.
                string dir = Path.GetDirectoryName(desiredPath) ?? string.Empty;
                string baseName = Path.GetFileNameWithoutExtension(desiredPath);
                string ext = Path.GetExtension(desiredPath);

                for (int i = 1; i <= 100; i++)
                {
                    string candidate = Path.Combine(dir, $"{baseName}_v{i}{ext}");
                    if (!File.Exists(candidate)) return candidate;

                    try
                    {
                        using var stream = File.Open(candidate, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                        return candidate;
                    }
                    catch (IOException) { }
                }

                return Path.Combine(dir, $"{baseName}_{DateTime.Now:yyyyMMdd_HHmmss}{ext}");
            }
        }
    }
}
