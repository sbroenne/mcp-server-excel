namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result of converting a worksheet range to an Excel table.
/// </summary>
public sealed class TableRangeConversionResult : ResultBase
{
    /// <summary>
    /// Worksheet containing the source range.
    /// </summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>
    /// Requested table name.
    /// </summary>
    public string TableName { get; set; } = string.Empty;

    /// <summary>
    /// Caller-supplied range.
    /// </summary>
    public string RequestedRange { get; set; } = string.Empty;

    /// <summary>
    /// Absolute range used after effective-range resolution.
    /// </summary>
    public string EffectiveRange { get; set; } = string.Empty;

    /// <summary>
    /// Merged-header policy used.
    /// </summary>
    public TableMergedHeaderPolicy MergedHeaderPolicy { get; set; }

    /// <summary>
    /// Header policy used.
    /// </summary>
    public TableHeaderPolicy HeaderPolicy { get; set; }

    /// <summary>
    /// Preflight findings observed before policy handling.
    /// </summary>
    public List<TablePreflightFinding> PreflightFindings { get; set; } = [];

    /// <summary>
    /// Merged header ranges normalized by the operation.
    /// </summary>
    public List<string> NormalizedMergedRanges { get; set; } = [];

    /// <summary>
    /// Header changes made by the operation.
    /// </summary>
    public List<TableHeaderChange> HeaderChanges { get; set; } = [];

    /// <summary>
    /// Created table metadata.
    /// </summary>
    public TableInfo? Table { get; set; }

    /// <summary>
    /// Post-creation validation details.
    /// </summary>
    public TableConversionValidationResult Validation { get; set; } = new();

    /// <summary>
    /// Rollback status. Successful conversions do not require rollback.
    /// </summary>
    public TableRollbackResult Rollback { get; set; } = new();
}
