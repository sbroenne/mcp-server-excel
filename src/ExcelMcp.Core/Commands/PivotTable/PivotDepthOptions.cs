using Sbroenne.ExcelMcp.Core.Models;
using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>Native PivotFilters calculation and comparison types.</summary>
public enum PivotFilterType
{
    /// <summary>Highest item counts.</summary>
    TopCount,
    /// <summary>Lowest item counts.</summary>
    BottomCount,
    /// <summary>Highest percentage.</summary>
    TopPercent,
    /// <summary>Lowest percentage.</summary>
    BottomPercent,
    /// <summary>Highest cumulative sum.</summary>
    TopSum,
    /// <summary>Lowest cumulative sum.</summary>
    BottomSum,
    /// <summary>Equal value.</summary>
    ValueEquals,
    /// <summary>Unequal value.</summary>
    ValueDoesNotEqual,
    /// <summary>Greater value.</summary>
    ValueIsGreaterThan,
    /// <summary>Greater or equal value.</summary>
    ValueIsGreaterThanOrEqualTo,
    /// <summary>Lower value.</summary>
    ValueIsLessThan,
    /// <summary>Lower or equal value.</summary>
    ValueIsLessThanOrEqualTo,
    /// <summary>Inclusive value interval.</summary>
    ValueIsBetween,
    /// <summary>Outside value interval.</summary>
    ValueIsNotBetween,
    /// <summary>Equal label.</summary>
    CaptionEquals,
    /// <summary>Unequal label.</summary>
    CaptionDoesNotEqual,
    /// <summary>Label prefix.</summary>
    CaptionBeginsWith,
    /// <summary>Excluded prefix.</summary>
    CaptionDoesNotBeginWith,
    /// <summary>Label suffix.</summary>
    CaptionEndsWith,
    /// <summary>Excluded suffix.</summary>
    CaptionDoesNotEndWith,
    /// <summary>Label substring.</summary>
    CaptionContains,
    /// <summary>Excluded substring.</summary>
    CaptionDoesNotContain,
    /// <summary>Label comparison.</summary>
    CaptionIsGreaterThan,
    /// <summary>Inclusive label comparison.</summary>
    CaptionIsGreaterThanOrEqualTo,
    /// <summary>Label comparison.</summary>
    CaptionIsLessThan,
    /// <summary>Inclusive label comparison.</summary>
    CaptionIsLessThanOrEqualTo,
    /// <summary>Label interval.</summary>
    CaptionIsBetween,
    /// <summary>Excluded label interval.</summary>
    CaptionIsNotBetween,
    /// <summary>Exact date.</summary>
    SpecificDate,
    /// <summary>Excluded date.</summary>
    NotSpecificDate,
    /// <summary>Earlier date.</summary>
    Before,
    /// <summary>Inclusive earlier date.</summary>
    BeforeOrEqualTo,
    /// <summary>Later date.</summary>
    After,
    /// <summary>Inclusive later date.</summary>
    AfterOrEqualTo,
    /// <summary>Inclusive date interval.</summary>
    DateBetween,
    /// <summary>Excluded date interval.</summary>
    DateNotBetween
}

/// <summary>Typed filter criteria; supply only the criteria required by type.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class PivotFilterOptions
{
    /// <summary>Native label, value, date, or top/bottom filter.</summary>
    [JsonRequired]
    public PivotFilterType Type { get; set; }
    /// <summary>Exact value-field caption for value and top/bottom filters.</summary>
    public string? DataFieldName { get; set; }
    /// <summary>First label criterion.</summary>
    public string? Text1 { get; set; }
    /// <summary>Second label criterion for interval filters.</summary>
    public string? Text2 { get; set; }
    /// <summary>First numeric criterion or top/bottom amount.</summary>
    public double? Number1 { get; set; }
    /// <summary>Second numeric criterion for interval filters.</summary>
    public double? Number2 { get; set; }
    /// <summary>First date criterion, ISO date/time in JSON.</summary>
    public DateTime? Date1 { get; set; }
    /// <summary>Second date criterion for interval filters.</summary>
    public DateTime? Date2 { get; set; }
}

/// <summary>Native PivotFilters entry, including criteria for filters created outside this API.</summary>
public sealed class PivotFilterInfo
{
    /// <summary>Current native collection index, not a persistent identifier.</summary>
    public int Index { get; set; }
    /// <summary>Native filter type, without the xl prefix.</summary>
    public string Type { get; set; } = string.Empty;
    /// <summary>Native first criterion where the type has one.</summary>
    public object? Value1 { get; set; }
    /// <summary>Native second criterion for interval types.</summary>
    public object? Value2 { get; set; }
    /// <summary>Native value-field caption for value/top filters.</summary>
    public string? DataFieldName { get; set; }
}

/// <summary>Complete calculated filters on the selected field; item visibility is separate.</summary>
public sealed class PivotFiltersResult : OperationResult
{
    /// <summary>Selected PivotTable.</summary>
    public string PivotTableName { get; set; } = string.Empty;
    /// <summary>Selected placed field.</summary>
    public string FieldName { get; set; } = string.Empty;
    /// <summary>All native PivotFilters entries.</summary>
    public List<PivotFilterInfo> Filters { get; set; } = [];
}

/// <summary>Omitted layout options retain native state.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class PivotLayoutOptions
{
    /// <summary>0=Compact, 1=Tabular, 2=Outline.</summary>
    public int? RowLayout { get; set; }
    /// <summary>Repeat all current row-field labels; requires Tabular or Outline layout.</summary>
    public bool? RepeatLabels { get; set; }
    /// <summary>Existing built-in or custom PivotTable style name.</summary>
    public string? StyleName { get; set; }
    /// <summary>Keep formatting on refresh.</summary>
    public bool? PreserveFormatting { get; set; }
    /// <summary>Native row banding.</summary>
    public bool? ShowRowStripes { get; set; }
    /// <summary>Native column banding.</summary>
    public bool? ShowColumnStripes { get; set; }
    /// <summary>Native row headers.</summary>
    public bool? ShowRowHeaders { get; set; }
    /// <summary>Native column headers.</summary>
    public bool? ShowColumnHeaders { get; set; }
    /// <summary>Permit multiple native calculated filters on each field.</summary>
    public bool? AllowMultipleFilters { get; set; }
}

/// <summary>Per-row-field native layout, including mixed configurations.</summary>
public sealed class PivotRowLayoutInfo
{
    /// <summary>Exact current field caption.</summary>
    public string FieldName { get; set; } = string.Empty;
    /// <summary>0=Compact, 1=Tabular, 2=Outline.</summary>
    public int RowLayout { get; set; }
    /// <summary>Native repeated-label setting.</summary>
    public bool RepeatLabels { get; set; }
}

/// <summary>Complete native PivotTable style and row layout.</summary>
public sealed class PivotLayoutResult : OperationResult
{
    /// <summary>Selected PivotTable.</summary>
    public string PivotTableName { get; set; } = string.Empty;
    /// <summary>Current native style name.</summary>
    public string StyleName { get; set; } = string.Empty;
    /// <summary>Native preserve-formatting setting.</summary>
    public bool PreserveFormatting { get; set; }
    /// <summary>Native row banding.</summary>
    public bool ShowRowStripes { get; set; }
    /// <summary>Native column banding.</summary>
    public bool ShowColumnStripes { get; set; }
    /// <summary>Native row headers.</summary>
    public bool ShowRowHeaders { get; set; }
    /// <summary>Native column headers.</summary>
    public bool ShowColumnHeaders { get; set; }
    /// <summary>Native permission for multiple calculated filters on each field.</summary>
    public bool AllowMultipleFilters { get; set; }
    /// <summary>Every placed row field.</summary>
    public List<PivotRowLayoutInfo> RowFields { get; set; } = [];
}

/// <summary>Selected native item expansion.</summary>
public sealed class PivotItemExpansionResult : OperationResult
{
    /// <summary>Selected PivotTable.</summary>
    public string PivotTableName { get; set; } = string.Empty;
    /// <summary>Selected field.</summary>
    public string FieldName { get; set; } = string.Empty;
    /// <summary>Selected exact item caption.</summary>
    public string ItemName { get; set; } = string.Empty;
    /// <summary>Native ShowDetail state.</summary>
    public bool Expanded { get; set; }
}

/// <summary>Native worksheet source/cache ownership.</summary>
public sealed class PivotSourceResult : OperationResult
{
    /// <summary>Selected PivotTable.</summary>
    public string PivotTableName { get; set; } = string.Empty;
    /// <summary>Current cache index.</summary>
    public int CacheIndex { get; set; }
    /// <summary>Native source reference for worksheet-backed PivotTables.</summary>
    public string SourceData { get; set; } = string.Empty;
    /// <summary>Native cache record count.</summary>
    public int RecordCount { get; set; }
    /// <summary>All PivotTables sharing the cache, including the selected table.</summary>
    public List<string> SharedPivotTables { get; set; } = [];
    /// <summary>All slicer/timeline caches connected to the selected table.</summary>
    public List<string> ConnectedSlicerCaches { get; set; } = [];
}
