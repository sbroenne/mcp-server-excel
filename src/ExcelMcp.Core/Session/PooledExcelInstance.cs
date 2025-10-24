namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Represents a pooled Excel COM instance with its associated workbook and metadata.
/// </summary>
internal class PooledExcelInstance
{
    public dynamic? Excel { get; set; }
    public dynamic? Workbook { get; set; }
    public DateTime LastUsed { get; set; }
    public required object Lock { get; set; }
}
