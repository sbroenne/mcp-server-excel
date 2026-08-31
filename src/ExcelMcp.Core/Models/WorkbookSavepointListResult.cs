namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result returned when listing a session's workbook savepoints.
/// </summary>
public sealed class WorkbookSavepointListResult : ResultBase
{
    /// <summary>Session that owns the savepoints.</summary>
    public string SessionId { get; set; } = string.Empty;

    /// <summary>Savepoints ordered by creation time.</summary>
    public List<WorkbookSavepointInfo> Savepoints { get; set; } = [];

    /// <summary>Number of retained savepoints.</summary>
    public int Count { get; set; }

    /// <summary>Total retained snapshot bytes for this session.</summary>
    public long TotalSizeBytes { get; set; }

    /// <summary>Maximum retained savepoints per session.</summary>
    public int MaxSavepointsPerSession { get; set; }

    /// <summary>Maximum retained snapshot bytes per session.</summary>
    public long MaxBytesPerSession { get; set; }

    /// <summary>Maximum retained snapshot bytes across the service process.</summary>
    public long MaxBytesPerProcess { get; set; }
}
