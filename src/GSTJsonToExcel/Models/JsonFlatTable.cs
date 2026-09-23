using System.Collections.Generic;

namespace GSTJsonToExcel.Models
{
    /// <summary>
    /// Represents a tabular structure destined for an Excel worksheet.
    /// </summary>
    public class JsonFlatTable
    {
        public string SheetName { get; set; } = string.Empty;
        public string FullJsonPath { get; set; } = string.Empty;
        public List<string> ColumnOrder { get; } = new();
        public HashSet<string> ColumnSet { get; } = new(System.StringComparer.OrdinalIgnoreCase);
        public List<JsonFlatRow> Rows { get; } = new();

        public void AddColumn(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName)) return;

            if (ColumnSet.Add(columnName))
            {
                ColumnOrder.Add(columnName);
            }
        }

        public void AddRow(JsonFlatRow row)
        {
            foreach (var col in row.Cells.Keys)
            {
                AddColumn(col);
            }
            Rows.Add(row);
        }

        public int RowCount => Rows.Count;
        public int ColumnCount => ColumnOrder.Count;
    }
}
