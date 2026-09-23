using System.Collections.Generic;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IDuplicateDetectorService
    {
        Task<string> ComputeFileHashAsync(string filePath);
        Task DetectDuplicatesAsync(List<ScannedFileItem> candidateFiles);
    }
}
