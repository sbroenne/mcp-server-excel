namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>Verified outcome of a query operation. Null row counts mean unavailable or not loaded.</summary>
public class PowerQueryOperationResult : OperationResult
{
    /// <summary>Name of the affected query.</summary>
    public string QueryName { get; set; } = string.Empty;

    /// <summary>Elapsed operation milliseconds, including queueing and destination inspection, excluding remote M formatting.</summary>
    public double DurationMs { get; set; }

    /// <summary>Loaded rows, without double-counting both destinations. Null when unavailable or destination counts disagree; zero for an explicit no-load operation.</summary>
    public long? RowsLoaded { get; set; }

    /// <summary>Verified worksheet table row count, excluding headers and totals; null if not present.</summary>
    public long? WorksheetRowsLoaded { get; set; }

    /// <summary>Verified model table record count; null if not present.</summary>
    public long? ModelRowsLoaded { get; set; }
}
