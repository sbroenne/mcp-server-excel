namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Rollback state for a range-to-table conversion.
/// </summary>
public sealed class TableRollbackResult
{
    /// <summary>
    /// Whether the operation changed workbook state and therefore required rollback.
    /// </summary>
    public bool Required { get; set; }

    /// <summary>
    /// Whether rollback was attempted.
    /// </summary>
    public bool Attempted { get; set; }

    /// <summary>
    /// Whether all rollback steps completed.
    /// </summary>
    public bool Completed { get; set; }

    /// <summary>
    /// Whether the restored state matched the captured rollback invariants.
    /// </summary>
    public bool Verified { get; set; }

    /// <summary>
    /// Rollback failure detail, when rollback did not complete or verify.
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    /// Very-hidden recovery worksheet retained when rollback could not be verified.
    /// </summary>
    public string? RecoverySheetName { get; set; }
}
