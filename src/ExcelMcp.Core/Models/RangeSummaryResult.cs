using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for per-column range summary statistics.
/// </summary>
public sealed class RangeSummaryResult : ResultBase
{
    /// <summary>Worksheet containing the source range.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Resolved source range address.</summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>Total rows in the source range.</summary>
    public int TotalRowCount { get; set; }

    /// <summary>Total columns in the source range.</summary>
    public int TotalColumnCount { get; set; }

    /// <summary>Selected worksheet columns in return order, or null when all columns are summarized.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? SelectedColumns { get; set; }

    /// <summary>Per-column statistics in requested order.</summary>
    public List<RangeColumnSummary> Columns { get; set; } = [];
}
