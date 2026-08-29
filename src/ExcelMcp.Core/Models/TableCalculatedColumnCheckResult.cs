namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Formula consistency check for one table column.
/// </summary>
public sealed class TableCalculatedColumnCheckResult
{
    /// <summary>
    /// Table column name.
    /// </summary>
    public string ColumnName { get; set; } = string.Empty;

    /// <summary>
    /// Uniform R1C1 formula pattern captured before sorting.
    /// </summary>
    public string FormulaR1C1 { get; set; } = string.Empty;

    /// <summary>
    /// Whether the column had one formula pattern before sorting.
    /// </summary>
    public bool ConsistentBefore { get; set; }

    /// <summary>
    /// Whether the same formula pattern remained after sorting.
    /// </summary>
    public bool ConsistentAfter { get; set; }

    /// <summary>
    /// Whether the calculated-column check passed.
    /// </summary>
    public bool Passed { get; set; }
}
