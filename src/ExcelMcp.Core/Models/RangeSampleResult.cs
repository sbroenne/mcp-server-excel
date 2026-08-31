using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for deterministic first/last row sampling.
/// </summary>
public sealed class RangeSampleResult : ResultBase
{
    /// <summary>Worksheet containing the source range.</summary>
    public string SheetName { get; set; } = string.Empty;

    /// <summary>Resolved source range address.</summary>
    public string RangeAddress { get; set; } = string.Empty;

    /// <summary>Total rows in the source range.</summary>
    public int TotalRowCount { get; set; }

    /// <summary>Total columns in the source range.</summary>
    public int TotalColumnCount { get; set; }

    /// <summary>Requested number of leading rows.</summary>
    public int FirstRowCount { get; set; }

    /// <summary>Requested number of trailing rows.</summary>
    public int LastRowCount { get; set; }

    /// <summary>Number of values returned for each sampled row.</summary>
    public int ColumnCount { get; set; }

    /// <summary>Selected worksheet columns in return order, or null when all columns are returned.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<string>? SelectedColumns { get; set; }

    /// <summary>Sampled rows in ascending source order without overlap duplicates.</summary>
    public List<RangeSampleRow> Rows { get; set; } = [];
}
