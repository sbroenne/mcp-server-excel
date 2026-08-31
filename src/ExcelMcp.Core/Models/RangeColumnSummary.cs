using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Type counts and numeric statistics for one worksheet column.
/// </summary>
public sealed class RangeColumnSummary
{
    /// <summary>Absolute worksheet column letters.</summary>
    public string Column { get; set; } = string.Empty;

    /// <summary>One-based absolute worksheet column number.</summary>
    public int ColumnNumber { get; set; }

    /// <summary>Resolved source address summarized for this column.</summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>Total cells in the summarized column range.</summary>
    public long CellCount { get; set; }

    /// <summary>Cells with neither a constant nor a formula.</summary>
    public long BlankCount { get; set; }

    /// <summary>Numeric cells, including Excel date serial values.</summary>
    public long NumericCount { get; set; }

    /// <summary>Text cells, including formulas returning text.</summary>
    public long TextCount { get; set; }

    /// <summary>Boolean cells.</summary>
    public long LogicalCount { get; set; }

    /// <summary>Constant or formula cells containing Excel errors.</summary>
    public long ErrorCount { get; set; }

    /// <summary>Sum of numeric cells, or null when no numeric cells exist.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Sum { get; set; }

    /// <summary>Average of numeric cells, or null when no numeric cells exist.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Average { get; set; }

    /// <summary>Minimum numeric value, or null when no numeric cells exist.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Minimum { get; set; }

    /// <summary>Maximum numeric value, or null when no numeric cells exist.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public double? Maximum { get; set; }
}
