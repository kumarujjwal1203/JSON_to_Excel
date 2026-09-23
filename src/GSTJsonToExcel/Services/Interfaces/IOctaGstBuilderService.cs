using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IOctaGstBuilderService
    {
        Task<(bool Success, int TotalRecords, string? Error)> BuildOctaWorkbookAsync(
            GstFileType fileType,
            List<ScannedFileItem> sourceFiles,
            string outputExcelPath,
            CancellationToken cancellationToken = default);
    }
}
