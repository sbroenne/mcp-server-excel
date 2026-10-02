using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>Native additional calculations, independent of the aggregation function.</summary>
public enum PivotFieldCalculation
{
    /// <summary>No additional calculation.</summary>
    Normal = -4143,
    /// <summary>Difference from a selected base item.</summary>
    DifferenceFrom = 2,
    /// <summary>Percentage of a selected base item.</summary>
    PercentOf = 3,
    /// <summary>Percentage difference from a selected base item.</summary>
    PercentDifferenceFrom = 4,
    /// <summary>Running total along the base field.</summary>
    RunningTotal = 5,
    /// <summary>Percentage of the row total.</summary>
    PercentOfRow = 6,
    /// <summary>Percentage of the column total.</summary>
    PercentOfColumn = 7,
    /// <summary>Percentage of the grand total.</summary>
    PercentOfTotal = 8,
    /// <summary>Native index calculation.</summary>
    Index = 9,
    /// <summary>Percentage of the parent row total.</summary>
    PercentOfParentRow = 10,
    /// <summary>Percentage of the parent column total.</summary>
    PercentOfParentColumn = 11,
    /// <summary>Percentage of the selected parent field total.</summary>
    PercentOfParent = 12,
    /// <summary>Running percentage along the base field.</summary>
    PercentRunningTotal = 13,
    /// <summary>Rank from smallest to largest along the base field.</summary>
    RankAscending = 14,
    /// <summary>Rank from largest to smallest along the base field.</summary>
    RankDescending = 15
}

/// <summary>Explicit base-item selection; names are never interpreted as indexes.</summary>
public enum PivotCalculationBaseItemKind
{
    /// <summary>Use an exact native item name.</summary>
    Named,
    /// <summary>Use the previous item in native PivotTable order.</summary>
    Previous,
    /// <summary>Use the next item in native PivotTable order.</summary>
    Next
}

/// <summary>Native settings for one displayed instance in the Values area.</summary>
public sealed class PivotFieldCalculationResult : ResultBase
{
    /// <summary>Exact displayed value-field name, used to select this instance.</summary>
    public string FieldName { get; set; } = string.Empty;
    /// <summary>Source field name; multiple displayed instances may share it.</summary>
    public string SourceName { get; set; } = string.Empty;
    /// <summary>Native position within Values.</summary>
    public int Position { get; set; }
    /// <summary>Native aggregation; OLAP measures define aggregation at the source.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public AggregationFunction? Function { get; set; }
    /// <summary>Native additional calculation.</summary>
    public PivotFieldCalculation Calculation { get; set; }
    /// <summary>Whether the PivotTable uses an OLAP/Data Model cache.</summary>
    public bool IsOlap { get; set; }
    /// <summary>Whether the native base-field/item properties are available.</summary>
    public bool BaseSettingsAvailable { get; set; }
    /// <summary>Applicable base field; absent when the calculation does not use one.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BaseFieldName { get; set; }
    /// <summary>Applicable native base-item kind.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public PivotCalculationBaseItemKind? BaseItemKind { get; set; }
    /// <summary>Exact native named base item, when applicable.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BaseItemName { get; set; }
}
