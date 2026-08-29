namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Structured failure details for range-to-table conversion.
/// </summary>
public sealed class TableRangeConversionFailureDetails
{
    /// <summary>
    /// Stage that failed.
    /// </summary>
    public TableConversionFailureStage FailureStage { get; set; }

    /// <summary>
    /// Whether caller cancellation initiated the failure and rollback.
    /// </summary>
    public bool WasCancelled { get; set; }

    /// <summary>
    /// Whether the session operation timeout initiated the failure and rollback.
    /// </summary>
    public bool WasTimedOut { get; set; }

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
    /// Resolved effective range, when available.
    /// </summary>
    public string? EffectiveRange { get; set; }

    /// <summary>
    /// Preflight findings observed before the failure.
    /// </summary>
    public List<TablePreflightFinding> PreflightFindings { get; set; } = [];

    /// <summary>
    /// Header changes made before the failure.
    /// </summary>
    public List<TableHeaderChange> HeaderChanges { get; set; } = [];

    /// <summary>
    /// Post-creation validation details, when validation ran.
    /// </summary>
    public TableConversionValidationResult? Validation { get; set; }

    /// <summary>
    /// Rollback status.
    /// </summary>
    public TableRollbackResult Rollback { get; set; } = new();
}
