using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using GSTJsonToExcel.Services.Interfaces;
using Microsoft.Win32;

namespace GSTJsonToExcel.Services.Implementations
{
    public class FileService : IFileService
    {
        public string? SelectJsonFile()
        {
            var files = SelectJsonFiles();
            return files != null && files.Length > 0 ? files[0] : null;
        }

        public string[]? SelectJsonFiles()
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select GST JSON / ZIP File(s)",
                Filter = "GST Files (*.json;*.zip)|*.json;*.zip|JSON Files (*.json)|*.json|ZIP Files (*.zip)|*.zip|All Files (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = true
            };

            return dialog.ShowDialog() == true ? dialog.FileNames : null;
        }

        public string? SelectFolder()
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Select Folder Containing GST Files",
                Multiselect = false
            };

            return dialog.ShowDialog() == true ? dialog.FolderName : null;
        }

        public string? SelectSaveLocation(string defaultFileName)
        {
            var dialog = new SaveFileDialog
            {
                Title = "Save Excel File As",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = defaultFileName,
                DefaultExt = ".xlsx",
                AddExtension = true
            };

            return dialog.ShowDialog() == true ? dialog.FileName : null;
        }

        public void OpenFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return;

            Process.Start(new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            });
        }

        public void OpenFolder(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            string? directory = File.Exists(filePath) ? Path.GetDirectoryName(filePath) : filePath;
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory)) return;

            if (File.Exists(filePath))
            {
                Process.Start("explorer.exe", $"/select,\"{filePath}\"");
            }
            else
            {
                Process.Start("explorer.exe", $"\"{directory}\"");
            }
        }

        public (bool IsValid, string? ErrorMessage, long FileSizeBytes) ValidateJsonFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return (false, "File does not exist.", 0);
            }

            try
            {
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.Length == 0)
                {
                    return (false, "The selected file is empty (0 bytes).", 0);
                }

                // Attempt to read and parse JSON syntax
                using var stream = File.OpenRead(filePath);
                using var document = JsonDocument.Parse(stream, new JsonDocumentOptions
                {
                    AllowTrailingCommas = true,
                    CommentHandling = JsonCommentHandling.Skip
                });

                return (true, null, fileInfo.Length);
            }
            catch (JsonException jEx)
            {
                return (false, $"Invalid JSON syntax: {jEx.Message}", 0);
            }
            catch (UnauthorizedAccessException)
            {
                return (false, "Access denied: The file could not be read due to permissions.", 0);
            }
            catch (IOException ioEx)
            {
                return (false, $"File I/O error (file may be locked): {ioEx.Message}", 0);
            }
            catch (Exception ex)
            {
                return (false, $"Unexpected validation error: {ex.Message}", 0);
            }
        }
    }
}
