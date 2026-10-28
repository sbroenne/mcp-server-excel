using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Filtering;

/// <summary>Native AutoFilter operator, including ordinary single-condition filtering.</summary>
public enum FilterOperator
{
    /// <summary>One comparison criterion.</summary>
    Comparison = 0,
    /// <summary>Both comparison criteria must match.</summary>
    And = 1,
    /// <summary>Either comparison criterion must match.</summary>
    Or = 2,
    /// <summary>Highest item count.</summary>
    TopItems = 3,
    /// <summary>Lowest item count.</summary>
    BottomItems = 4,
    /// <summary>Highest percentage.</summary>
    TopPercent = 5,
    /// <summary>Lowest percentage.</summary>
    BottomPercent = 6,
    /// <summary>Explicit values or date groups.</summary>
    Values = 7,
    /// <summary>Native displayed cell color.</summary>
    CellColor = 8,
    /// <summary>Native displayed font color.</summary>
    FontColor = 9,
    /// <summary>Native conditional-format icon.</summary>
    Icon = 10,
    /// <summary>Native relative date or average criterion.</summary>
    Dynamic = 11,
    /// <summary>Cells without fill.</summary>
    NoFill = 12,
    /// <summary>Cells without font color.</summary>
    AutomaticFontColor = 13,
    /// <summary>Cells without an icon.</summary>
    NoIcon = 14
}

/// <summary>Native date-group granularity.</summary>
public enum FilterDateLevel
{
    /// <summary>Whole year containing Date.</summary>
    Year = 0,
    /// <summary>Whole month containing Date.</summary>
    Month = 1,
    /// <summary>Exact day.</summary>
    Day = 2
}

/// <summary>One native date group.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class FilterDateGroup
{
    /// <summary>Year, Month, or Day.</summary>
    public FilterDateLevel Level { get; set; }
    /// <summary>Calendar date identifying the native group.</summary>
    public DateTime Date { get; set; }
}

/// <summary>Typed native filtering settings shared by tables and ordinary ranges.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class FilterOptions
{
    /// <summary>Native operator; Comparison by default.</summary>
    public FilterOperator FilterOperator { get; set; }
    /// <summary>First native comparison string, such as greater-than 100 or equals Active.</summary>
    public string? Criteria1 { get; set; }
    /// <summary>Second native comparison; required only for And or Or.</summary>
    public string? Criteria2 { get; set; }
    /// <summary>Exact displayed values for Values filtering; cannot combine with dateGroups.</summary>
    public List<string>? Values { get; set; }
    /// <summary>Native date groups for Values filtering; cannot combine with values.</summary>
    public List<FilterDateGroup>? DateGroups { get; set; }
    /// <summary>Positive whole item count or percentage for top/bottom operators.</summary>
    public int? Count { get; set; }
    /// <summary>RGB hex color for CellColor or FontColor.</summary>
    public string? Color { get; set; }
    /// <summary>Native XlIconSet name for Icon filtering.</summary>
    public string? IconSet { get; set; }
    /// <summary>One-based icon index within the selected set.</summary>
    public int? IconIndex { get; set; }
    /// <summary>Native XlDynamicFilterCriteria name for Dynamic filtering.</summary>
    public string? DynamicCriteria { get; set; }
}

/// <summary>One native criterion getter, with explicit failure coverage instead of an invented empty criterion.</summary>
public sealed class FilterCriterion
{
    /// <summary>Whether Excel returned the criterion.</summary>
    public bool Available { get; set; }
    /// <summary>Whether Excel returned an empty native variant rather than a populated criterion.</summary>
    public bool EmptyVariant { get; set; }
    /// <summary>Native scalar, ordered array, or icon descriptor; never an array's ToString output.</summary>
    public object? Value { get; set; }
    /// <summary>Native getter error when the criterion could not be read.</summary>
    public string? ReadError { get; set; }
    /// <summary>Native HRESULT when a criterion getter failed.</summary>
    public int? HResult { get; set; }
}

/// <summary>Native conditional-format icon identity.</summary>
public sealed class FilterIcon
{
    /// <summary>Native XlIconSet name.</summary>
    public string IconSet { get; set; } = string.Empty;
    /// <summary>One-based native icon index.</summary>
    public int IconIndex { get; set; }
}
