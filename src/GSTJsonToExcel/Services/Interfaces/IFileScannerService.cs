using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IFileScannerService
    {
        Task<BatchScanSummary> ScanPathsAsync(
            IEnumerable<string> pathsOrDirectories,
            CancellationToken cancellationToken = default);
    }
}
