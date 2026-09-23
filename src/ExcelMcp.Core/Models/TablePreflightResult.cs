namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Non-destructive safety report for creating an Excel table from a range.
/// </summary>
public sealed class TablePreflightResult : ResultBase
{
    /// <summary>
    /// Worksheet containing the proposed table.
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Requested table name.
    /// </summary>
    public string TableName { get; set; } = string.Empty;

    /// <summary>
    /// Range supplied by the caller.
    /// </summary>
    public string RequestedRange { get; set; } = string.Empty;

    /// <summary>
    /// Absolute range Excel will use after single-cell CurrentRegion expansion.
    /// </summary>
    public string EffectiveRange { get; set; } = string.Empty;

    /// <summary>
    /// True when no deterministic blocker was found.
    /// </summary>
    public bool SafeToCreate { get; set; }

    /// <summary>
    /// Blocking and advisory findings for the proposed table.
    /// </summary>
    public List<TablePreflightFinding> Findings { get; set; } = [];
}
