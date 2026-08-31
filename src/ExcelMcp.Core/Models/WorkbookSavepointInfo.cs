namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Public metadata for one immutable workbook savepoint.
/// </summary>
public sealed class WorkbookSavepointInfo
{
    /// <summary>Caller-provided savepoint name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>UTC creation time.</summary>
    public DateTime CreatedAtUtc { get; set; }

    /// <summary>Snapshot size in bytes.</summary>
    public long SizeBytes { get; set; }
}
