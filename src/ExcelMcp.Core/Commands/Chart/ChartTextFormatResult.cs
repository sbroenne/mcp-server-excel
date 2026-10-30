using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>Existing chart element whose text is selected.</summary>
public enum ChartTextElement
{
    /// <summary>Excel's whole-chart text setting; individual overrides may differ.</summary>
    ChartArea,
    /// <summary>Existing chart title.</summary>
    Title,
    /// <summary>Category-axis tick labels.</summary>
    CategoryAxis,
    /// <summary>Value-axis tick labels.</summary>
    ValueAxis,
    /// <summary>Existing series data labels.</summary>
    DataLabels,
    /// <summary>Existing chart legend.</summary>
    Legend,
    /// <summary>Title of the selected existing axis.</summary>
    AxisTitle
}

/// <summary>Actual native formatting of the selected chart element.</summary>
public sealed class ChartTextFormatResult : OperationResult
{
    /// <summary>Selected embedded chart name.</summary>
    public string ChartName { get; set; } = string.Empty;
    /// <summary>Selected text element.</summary>
    public ChartTextElement Element { get; set; }
    /// <summary>One entry per target; all-series data labels return one entry per series.</summary>
    public List<ChartTextFormat> Formats { get; set; } = [];
}

/// <summary>Native font properties, with explicit mixed fields.</summary>
public sealed class ChartTextFormat
{
    /// <summary>Selected axis, when applicable.</summary>
    public ChartAxisType? Axis { get; set; }
    /// <summary>One-based series index for data labels.</summary>
    public int? SeriesIndex { get; set; }
    /// <summary>Native typeface or theme-font token; null when mixed.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public string? FontName { get; set; }
    /// <summary>Native font size in points; null when mixed.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public double? FontSize { get; set; }
    /// <summary>Native bold setting; null when mixed.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public bool? Bold { get; set; }
    /// <summary>Native italic setting; null when mixed.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public bool? Italic { get; set; }
    /// <summary>Native text colour as #RRGGBB; null when mixed.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.Never)]
    public string? Color { get; set; }
    /// <summary>CamelCase names of properties with mixed native values.</summary>
    public List<string> MixedFields { get; set; } = [];
}
