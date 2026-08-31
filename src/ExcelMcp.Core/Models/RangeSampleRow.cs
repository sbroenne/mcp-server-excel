namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// A sampled source row with relative and absolute coordinates.
/// </summary>
public sealed class RangeSampleRow
{
    /// <summary>Zero-based row offset within the source range.</summary>
    public int RowOffset { get; set; }

    /// <summary>One-based absolute worksheet row number.</summary>
    public int RowNumber { get; set; }

    /// <summary>Absolute address of the returned cells in this row.</summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>Cell values in selected column order.</summary>
    public List<object?> Values { get; set; } = [];
}
