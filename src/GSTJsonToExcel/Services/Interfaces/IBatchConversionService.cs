using System;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IBatchConversionService
    {
        Task<BatchConversionResult> ConvertBatchAsync(
            BatchScanSummary scanSummary,
            string targetOutputDirectory,
            BatchConversionMode mode = BatchConversionMode.IndividualFiles,
            IProgress<BatchConversionProgress>? progress = null,
            CancellationToken cancellationToken = default);
    }
}
