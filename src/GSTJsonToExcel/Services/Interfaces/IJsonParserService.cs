using System;
using System.Threading;
using System.Threading.Tasks;
using GSTJsonToExcel.Models;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IJsonParserService
    {
        Task<JsonParseResult> ParseAndFlattenAsync(
            string jsonFilePath, 
            IProgress<ConversionProgress>? progress = null, 
            CancellationToken cancellationToken = default);
    }
}
