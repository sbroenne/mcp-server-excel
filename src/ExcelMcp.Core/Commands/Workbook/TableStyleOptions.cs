using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>Changes to a custom native table/Pivot/slicer/timeline style; omitted settings retain native values.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class TableStyleOptions
{
    /// <summary>Offer this style for worksheet tables.</summary>
    public bool? ShowAsAvailableTableStyle { get; set; }
    /// <summary>Offer this style for PivotTables.</summary>
    public bool? ShowAsAvailablePivotTableStyle { get; set; }
    /// <summary>Offer this style for slicers.</summary>
    public bool? ShowAsAvailableSlicerStyle { get; set; }
    /// <summary>Offer this style for timelines.</summary>
    public bool? ShowAsAvailableTimelineStyle { get; set; }
    /// <summary>Changes to unique native element types. Omitted elements retain their definitions.</summary>
    public List<TableStyleElementOptions>? Elements { get; set; }
}

/// <summary>Native differential formatting: no font name/size, script, alignment, number format, or diagonal borders.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class TableStyleElementOptions
{
    /// <summary>Exact native XlTableStyleElementType name, such as xlHeaderRow or xlRowStripe1.</summary>
    public required string ElementType { get; set; }
    /// <summary>Remove this element's entire native formatting. Cannot be combined with other settings.</summary>
    public bool? Clear { get; set; }
    /// <summary>Positive row/column stripe size; requires an existing format or formatting in the same element request.</summary>
    public int? StripeSize { get; set; }
    /// <summary>Font emphasis.</summary>
    public bool? Bold { get; set; }
    /// <summary>Italic font.</summary>
    public bool? Italic { get; set; }
    /// <summary>none, single, double, singleAccounting, or doubleAccounting.</summary>
    public string? Underline { get; set; }
    /// <summary>Strike through text.</summary>
    public bool? Strikethrough { get; set; }
    /// <summary>Theme font: 0 none, 1 major, 2 minor.</summary>
    public int? ThemeFont { get; set; }
    /// <summary>Fixed font RGB; mutually exclusive with fontThemeColor.</summary>
    public string? FontColor { get; set; }
    /// <summary>Native theme color 1 through 12.</summary>
    public int? FontThemeColor { get; set; }
    /// <summary>Font tint/shade between -1 and 1.</summary>
    public double? FontTintAndShade { get; set; }
    /// <summary>Solid fill RGB; mutually exclusive with fillThemeColor.</summary>
    public string? FillColor { get; set; }
    /// <summary>Native fill theme color 1 through 12.</summary>
    public int? FillThemeColor { get; set; }
    /// <summary>Fill tint/shade between -1 and 1.</summary>
    public double? FillTintAndShade { get; set; }
    /// <summary>Four outer and two inside borders; diagonals are not supported by native table-style elements.</summary>
    public List<CellBorderOptions>? Borders { get; set; }

    internal CellFormatOptions ToFormatOptions() => new()
    {
        Bold = Bold,
        Italic = Italic,
        Underline = Underline,
        Strikethrough = Strikethrough,
        ThemeFont = ThemeFont,
        FontColor = FontColor,
        FontThemeColor = FontThemeColor,
        FontTintAndShade = FontTintAndShade,
        FillColor = FillColor,
        FillThemeColor = FillThemeColor,
        FillTintAndShade = FillTintAndShade,
        Borders = Borders
    };

    internal bool HasFontFormatting => Bold is not null || Italic is not null || Underline is not null ||
        Strikethrough is not null || ThemeFont is not null || FontColor is not null || FontThemeColor is not null ||
        FontTintAndShade is not null;
    internal bool HasFillFormatting => FillColor is not null || FillThemeColor is not null || FillTintAndShade is not null;
    internal bool HasFormatting => HasFontFormatting || HasFillFormatting || Borders?.Count > 0;
}

/// <summary>Native style identity and supported user-interface availability.</summary>
public sealed record TableStyleInfo(string Name, string NameLocal, bool BuiltIn,
    bool ShowAsAvailableTableStyle, bool ShowAsAvailablePivotTableStyle,
    bool ShowAsAvailableSlicerStyle, bool ShowAsAvailableTimelineStyle);

/// <summary>One native differential element. Unformatted elements have no invented formatting or stripe size.</summary>
public sealed record TableStyleElementDefinition(string ElementType, int NativeType, bool HasFormat,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] int? StripeSize,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] CellFontFormat? Font,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] CellFillFormat? Fill,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.Never)] List<CellBorderFormat>? Borders,
    List<string> UnsetOrUnavailableFields);

/// <summary>Every native element of the selected style, including table/Pivot/slicer/timeline elements.</summary>
public sealed record TableStyleDefinition(string Name, string NameLocal, bool BuiltIn,
    bool ShowAsAvailableTableStyle, bool ShowAsAvailablePivotTableStyle,
    bool ShowAsAvailableSlicerStyle, bool ShowAsAvailableTimelineStyle,
    List<TableStyleElementDefinition> Elements);

/// <summary>All native workbook table-style identities, without a cap.</summary>
public sealed class TableStyleListResult : ResultBase
{
    /// <summary>Complete native style catalogue.</summary>
    public List<TableStyleInfo> Styles { get; init; } = [];
}

/// <summary>Complete native table-style definition and explicit inspection limitations.</summary>
public sealed class TableStyleResult : ResultBase
{
    /// <summary>Identity, availability, and all elements.</summary>
    public required TableStyleDefinition Style { get; init; }
    /// <summary>Native restrictions; unsupported state is never invented.</summary>
    public List<string> ReadLimitations { get; init; } =
        ["Table-style elements are differential formats: font name/size, subscript/superscript, alignment, number format and diagonal borders are not supported. Unset or unavailable properties remain null."];
}
