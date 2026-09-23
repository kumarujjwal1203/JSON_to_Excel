using System;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IExcelExportService
    {
        Task<ConversionResult> ExportToExcelAsync(
            JsonParseResult parseResult, 
            string outputFilePath, 
            IProgress<ConversionProgress>? progress = null, 
            CancellationToken cancellationToken = default);
    }
}
