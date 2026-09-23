using System.Collections.Generic;

namespace GSTJsonToExcel.Models
{
    public class JsonParseResult
    {
        public string SourceFileName { get; set; } = string.Empty;
        public string SourceFilePath { get; set; } = string.Empty;
        public long SourceFileSizeBytes { get; set; }
        public List<JsonFlatTable> Tables { get; } = new();
        public Dictionary<string, JsonFlatValue> AllLeafNodes { get; } = new(System.StringComparer.Ordinal);
        public List<string> EmptySections { get; } = new();

        public int TotalLeafCount => AllLeafNodes.Count;
    }
}
