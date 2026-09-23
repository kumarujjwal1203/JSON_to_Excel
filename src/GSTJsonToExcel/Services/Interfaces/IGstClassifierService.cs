using System.Threading.Tasks;
using GSTJsonToExcel.Models;

namespace GSTJsonToExcel.Services.Interfaces
{
    public interface IGstClassifierService
    {
        Task<(GstFileType Type, string Reason)> ClassifyFileAsync(string filePath);
        (GstFileType Type, string Reason) ClassifyByFileName(string fileName);
    }
}
