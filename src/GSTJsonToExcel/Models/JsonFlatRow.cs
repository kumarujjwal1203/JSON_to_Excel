using System.Collections.Generic;

namespace GSTJsonToExcel.Models
{
    /// <summary>
    /// Represents a single flat row of key-value cells in an exported worksheet.
    /// </summary>
    public class JsonFlatRow
    {
        public Dictionary<string, JsonFlatValue> Cells { get; } = new(System.StringComparer.OrdinalIgnoreCase);

        public string? ParentKey { get; set; }
        public string? SourceJsonPath { get; set; }

        public void Set(string column, JsonFlatValue value)
        {
            Cells[column] = value;
        }

        public JsonFlatValue? Get(string column)
        {
            return Cells.TryGetValue(column, out var val) ? val : null;
        }
    }
}
