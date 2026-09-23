using System.Collections.Generic;

namespace GSTJsonToExcel.Models
{
    public class DataIntegrityReport
    {
        public bool Passed { get; set; } = true;
        public int TotalSourceLeafValues { get; set; }
        public int TotalExportedCells { get; set; }
        public int MatchedLeafValues { get; set; }
        public double MatchPercentage => TotalSourceLeafValues > 0 
            ? (double)MatchedLeafValues / TotalSourceLeafValues * 100.0 
            : 100.0;

        public List<string> SectionsVerified { get; } = new();
        public List<string> Warnings { get; } = new();
        public List<string> AuditNotes { get; } = new();
    }
}
