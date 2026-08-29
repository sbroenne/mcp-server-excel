namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Before/after numeric control-total comparison.
/// </summary>
public sealed class TableControlTotalCheckResult
{
    /// <summary>
    /// Table column name.
    /// </summary>
    public string ColumnName { get; set; } = string.Empty;

    /// <summary>
    /// Numeric sum before sorting.
    /// </summary>
    public decimal Before { get; set; }

    /// <summary>
    /// Numeric sum after sorting.
    /// </summary>
    public decimal After { get; set; }

    /// <summary>
    /// Signed difference between the after and before sums.
    /// </summary>
    public decimal Delta { get; set; }

    /// <summary>
    /// Allowed absolute difference.
    /// </summary>
    public decimal Tolerance { get; set; }

    /// <summary>
    /// Whether the total stayed within tolerance.
    /// </summary>
    public bool Passed { get; set; }
}
