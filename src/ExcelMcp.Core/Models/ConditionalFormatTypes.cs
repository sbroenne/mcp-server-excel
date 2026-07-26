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
    /// Border style name, when set. Read by scanning the edge borders
    /// (left, top, bottom, right) and using the first one that has a style.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BorderStyle { get; set; }

    /// <summary>
    /// Border color as #RRGGBB, when set. Read from the same edge border as
    /// <see cref="BorderStyle"/> (the first edge that has a style).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BorderColor { get; set; }

    /// <summary>
    /// Color-scale criteria (2 or 3 stops), populated for <c>colorScale</c> rules.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<ColorScaleCriterionInfo>? ColorScaleCriteria { get; set; }

    /// <summary>
    /// Data-bar configuration, populated for <c>dataBar</c> rules.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public DataBarInfo? DataBar { get; set; }

    /// <summary>
    /// Icon-set configuration, populated for <c>iconSet</c> rules.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IconSetInfo? IconSet { get; set; }

    /// <summary>
    /// Top/bottom configuration, populated for <c>top10</c> rules.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public Top10Info? Top10 { get; set; }

    /// <summary>
    /// Above/below average selector (e.g. aboveAverage, belowAverage, aboveStdDev),
    /// populated for <c>aboveAverage</c> rules.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? AboveBelow { get; set; }

    /// <summary>
    /// Date period (e.g. today, yesterday, last7Days, thisMonth),
    /// populated for <c>timePeriod</c> rules.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DatePeriod { get; set; }
}

/// <summary>
/// A single color-scale stop (minimum, midpoint or maximum).
/// </summary>
public class ColorScaleCriterionInfo
{
    /// <summary>
    /// Threshold type (e.g. minimum, maximum, number, percent, percentile, formula).
    /// </summary>
    public string Type { get; set; } = "";

    /// <summary>
    /// Threshold value, when the type requires one (null for minimum/maximum/automatic).
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; set; }

    /// <summary>
    /// Stop color as #RRGGBB.
    /// </summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Color { get; set; }
}

/// <summary>
/// Data-bar configuration read from a <c>dataBar</c> rule.
/// </summary>
public class DataBarInfo
{
    /// <summary>Bar fill color as #RRGGBB.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FillColor { get; set; }

    /// <summary>Negative-value bar color as #RRGGBB, when a custom negative color is set.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BarColorNegative { get; set; }

    /// <summary>Fill direction (context, leftToRight, rightToLeft).</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Direction { get; set; }

    /// <summary>Whether the cell value is shown alongside the bar.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? ShowValue { get; set; }

    /// <summary>Minimum-point threshold type.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MinType { get; set; }

    /// <summary>Minimum-point threshold value, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MinValue { get; set; }

    /// <summary>Maximum-point threshold type.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MaxType { get; set; }

    /// <summary>Maximum-point threshold value, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MaxValue { get; set; }
}

/// <summary>
/// Icon-set configuration read from an <c>iconSet</c> rule.
/// </summary>
public class IconSetInfo
{
    /// <summary>Icon-set id (e.g. 3Arrows, 3TrafficLights1, 5Quarters).</summary>
    public string Id { get; set; } = "";

    /// <summary>Whether the icon order is reversed.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Reverse { get; set; }

    /// <summary>Whether only the icon (not the value) is shown.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? ShowIconOnly { get; set; }

    /// <summary>Icon threshold criteria, in ascending order.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<IconCriterionInfo>? Criteria { get; set; }
}

/// <summary>
/// A single icon-set threshold criterion.
/// </summary>
public class IconCriterionInfo
{
    /// <summary>Comparison operator (greater, greaterEqual).</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Operator { get; set; }

    /// <summary>Threshold value, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; set; }

    /// <summary>Threshold type (number, percent, percentile, formula).</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Type { get; set; }

    /// <summary>Icon name for this criterion (e.g. greenUpArrow), when readable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Icon { get; set; }
}

/// <summary>
/// Top/bottom configuration read from a <c>top10</c> rule.
/// </summary>
public class Top10Info
{
    /// <summary>Number of values (or percent) to highlight.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Rank { get; set; }

    /// <summary>Whether <see cref="Rank"/> is a percentage rather than a count.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Percent { get; set; }

    /// <summary>Whether the rule highlights the top or bottom values.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? TopBottom { get; set; }
}
