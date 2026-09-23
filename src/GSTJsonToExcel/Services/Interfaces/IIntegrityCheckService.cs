using System;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IIntegrityCheckService
    {
        Task<DataIntegrityReport> VerifyIntegrityAsync(
            JsonParseResult parseResult, 
            string excelFilePath, 
            IProgress<ConversionProgress>? progress = null, 
            CancellationToken cancellationToken = default);
    }
}
