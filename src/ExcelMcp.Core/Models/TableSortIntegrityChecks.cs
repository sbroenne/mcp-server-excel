namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Deterministic checks performed around a validated table sort.
/// </summary>
public sealed class TableSortIntegrityChecks
{
    /// <summary>
    /// Whether the table range address stayed unchanged.
    /// </summary>
    public bool? RangePreserved { get; set; }

    /// <summary>
    /// Whether the table row and column counts stayed unchanged.
    /// </summary>
    public bool? ShapePreserved { get; set; }

    /// <summary>
    /// Whether table headers stayed unchanged.
    /// </summary>
    public bool? HeadersPreserved { get; set; }

    /// <summary>
    /// Whether totals-row visibility and content stayed unchanged.
    /// </summary>
    public bool? TotalsRowPreserved { get; set; }

    /// <summary>
    /// Whether complete logical table rows were only permuted.
    /// </summary>
    public bool? RowSetPreserved { get; set; }

    /// <summary>
    /// Formula-pattern checks for calculated columns.
    /// </summary>
    public List<TableCalculatedColumnCheckResult> CalculatedColumns { get; set; } = [];

    /// <summary>
    /// Optional composite row-key check.
    /// </summary>
    public TableRowKeyCheckResult? RowKeys { get; set; }

    /// <summary>
    /// Optional numeric control-total checks.
    /// </summary>
    public List<TableControlTotalCheckResult> ControlTotals { get; set; } = [];
}
