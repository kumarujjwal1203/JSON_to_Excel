namespace GSTJsonToExcel.Models
{
    public enum GstFileType
    {
        R1,
        R3A,
        R2A,
        R2B,
        Unknown
    }

    public enum FileValidationStatus
    {
        Valid,
        Duplicate,
        InvalidJson,
        Empty,
        UnknownType,
        Ignored
    }

    public enum BatchConversionMode
    {
        IndividualFiles, // 1 Excel workbook per JSON file (Default)
        MergedByType     // 1 Excel workbook per GST return type
    }

    public enum ProcessingState
    {
        Pending,
        Converting,
        Completed,
        Failed,
        Skipped
    }
}
