namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result of sorting an Excel table, including optional integrity evidence.
/// </summary>
public sealed class TableSortResult : OperationResult
{
    /// <summary>
    /// Name of the sorted table.
    /// </summary>
    public string TableName { get; set; } = string.Empty;

    /// <summary>
    /// Absolute table range observed before sorting.
    /// </summary>
    public string TableRange { get; set; } = string.Empty;

    /// <summary>
    /// Whether integrity validation was performed.
    /// </summary>
    public bool ValidationPerformed { get; set; }

    /// <summary>
    /// Whether Excel's sort operation was invoked.
    /// </summary>
    public bool SortAttempted { get; set; }

    /// <summary>
    /// Whether the sorted state was kept.
    /// </summary>
    public bool SortCommitted { get; set; }

    /// <summary>
    /// Whether all requested integrity checks passed, or null when validation was not performed.
    /// </summary>
    public bool? IntegrityPreserved { get; set; }

    /// <summary>
    /// Whether restoration of the pre-sort snapshot was attempted.
    /// </summary>
    public bool RollbackAttempted { get; set; }

    /// <summary>
    /// Whether restoration was verified, or null when rollback was not attempted.
    /// </summary>
    public bool? RollbackSucceeded { get; set; }

    /// <summary>
    /// State covered by rollback verification. TableContent excludes row-specific formatting.
    /// </summary>
    public string RollbackScope { get; set; } = "TableContent";

    /// <summary>
    /// Blocking and advisory integrity findings.
    /// </summary>
    public List<TablePreflightFinding> Findings { get; set; } = [];

    /// <summary>
    /// Typed post-sort integrity checks.
    /// </summary>
    public TableSortIntegrityChecks Checks { get; set; } = new();
}
