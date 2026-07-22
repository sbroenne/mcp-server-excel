using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Models;

/// <summary>
/// Result for listing existing conditional formatting rules from a range or worksheet.
/// </summary>
public class ConditionalFormatListResult : ResultBase
{
    /// <summary>
    /// Worksheet the rules were read from.
    /// </summary>
    public string SheetName { get; set; } = "";

    /// <summary>
    /// Range the rules were read from. Null when the whole worksheet was scanned.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RangeAddress { get; set; }

    /// <summary>
    /// Conditional formatting rules in priority order.
    /// </summary>
    public List<ConditionalFormatRuleInfo> Rules { get; set; } = [];
}

/// <summary>
/// Describes a single conditional formatting rule read from Excel.
/// Formatting properties are only populated when the rule actually sets them.
/// </summary>
public class ConditionalFormatRuleInfo
{
    /// <summary>
    /// Rule type (e.g. cellValue, expression, colorScale, dataBar, top10, iconSet,
    /// uniqueValues, blanksCondition, timePeriod, aboveAverage).
    /// </summary>
    public string Type { get; set; } = "";

    /// <summary>
    /// Comparison operator (e.g. equal, greater, between). Null when the rule type
    /// does not use an operator (e.g. expression, colorScale).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Operator { get; set; }

    /// <summary>
    /// First formula/value for the condition.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula1 { get; set; }

    /// <summary>
    /// Second formula/value (used by between/notBetween).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Formula2 { get; set; }

    /// <summary>
    /// Range the rule applies to (absolute address, e.g. $A$1:$A$41).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AppliesTo { get; set; }

    /// <summary>
    /// Priority of the rule within the collection (1-based, lower = higher priority).
    /// Null when the rule type does not expose a priority.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Priority { get; set; }

    /// <summary>
    /// Whether rule evaluation stops if this rule is true.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? StopIfTrue { get; set; }

    /// <summary>
    /// Interior (fill) color as #RRGGBB, when set.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? InteriorColor { get; set; }

    /// <summary>
    /// Interior pattern constant, when set.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? InteriorPattern { get; set; }

    /// <summary>
    /// Font color as #RRGGBB, when set.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontColor { get; set; }

    /// <summary>
    /// Font bold state, when set (null = not specified / mixed).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? FontBold { get; set; }

    /// <summary>
    /// Font italic state, when set (null = not specified / mixed).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? FontItalic { get; set; }

    /// <summary>
    /// Border style name, when set (applied to edge borders).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BorderStyle { get; set; }

    /// <summary>
    /// Border color as #RRGGBB, when set.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BorderColor { get; set; }
}
