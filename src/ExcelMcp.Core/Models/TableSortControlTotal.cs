namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Optional numeric control total to compare before and after a table sort.
/// </summary>
public sealed class TableSortControlTotal
{
    /// <summary>
    /// Table column whose numeric values should be summed.
    /// </summary>
    public string ColumnName { get; set; } = string.Empty;

    /// <summary>
    /// Allowed absolute difference between the before and after sums.
    /// </summary>
    public decimal Tolerance { get; set; }
}
