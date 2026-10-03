using System.Text.Json.Serialization;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>Native Excel border positions.</summary>
public enum CellBorderPosition
{
    /// <summary>Descending diagonal.</summary>
    DiagonalDown = 5,
    /// <summary>Ascending diagonal.</summary>
    DiagonalUp = 6,
    /// <summary>Left edge.</summary>
    Left = 7,
    /// <summary>Top edge.</summary>
    Top = 8,
    /// <summary>Bottom edge.</summary>
    Bottom = 9,
    /// <summary>Right edge.</summary>
    Right = 10,
    /// <summary>Internal column boundaries.</summary>
    InsideVertical = 11,
    /// <summary>Internal row boundaries.</summary>
    InsideHorizontal = 12
}

/// <summary>One selected border. Omitted settings retain their native values.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class CellBorderOptions
{
    /// <summary>Required native border position.</summary>
    [JsonRequired]
    public CellBorderPosition Position { get; set; }
    /// <summary>continuous, dash, dot, dashdot, dashdotdot, double, slantdashdot, or none.</summary>
    public string? LineStyle { get; set; }
    /// <summary>hairline, thin, medium, or thick.</summary>
    public string? Weight { get; set; }
    /// <summary>Fixed #RRGGBB; mutually exclusive with themeColor.</summary>
    public string? Color { get; set; }
    /// <summary>Native workbook theme-color index 1 through 12.</summary>
    public int? ThemeColor { get; set; }
    /// <summary>Native tint/shade, -1 through 1.</summary>
    public double? TintAndShade { get; set; }
}

/// <summary>Shared visual formatting. Omitted properties preserve existing settings.</summary>
[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class CellFormatOptions
{
    /// <summary>Explicit font family; mutually exclusive with themeFont.</summary>
    public string? FontName { get; set; }
    /// <summary>Font size in points, 1 through 409.</summary>
    public double? FontSize { get; set; }
    /// <summary>Bold state.</summary>
    public bool? Bold { get; set; }
    /// <summary>Italic state.</summary>
    public bool? Italic { get; set; }
    /// <summary>none, single, double, singleAccounting, or doubleAccounting.</summary>
    public string? Underline { get; set; }
    /// <summary>Strikethrough state.</summary>
    public bool? Strikethrough { get; set; }
    /// <summary>Subscript state; cannot be true with superscript.</summary>
    public bool? Subscript { get; set; }
    /// <summary>Superscript state; cannot be true with subscript.</summary>
    public bool? Superscript { get; set; }
    /// <summary>Native theme font: 0 none, 1 major, 2 minor.</summary>
    public int? ThemeFont { get; set; }
    /// <summary>Fixed #RRGGBB; mutually exclusive with fontThemeColor.</summary>
    public string? FontColor { get; set; }
    /// <summary>Native theme-color index 1 through 12.</summary>
    public int? FontThemeColor { get; set; }
    /// <summary>Font tint/shade, -1 through 1.</summary>
    public double? FontTintAndShade { get; set; }
    /// <summary>Fixed #RRGGBB; mutually exclusive with fillThemeColor.</summary>
    public string? FillColor { get; set; }
    /// <summary>Native theme-color index 1 through 12.</summary>
    public int? FillThemeColor { get; set; }
    /// <summary>Fill tint/shade, -1 through 1.</summary>
    public double? FillTintAndShade { get; set; }
    /// <summary>Independent edge, inside, and diagonal border settings.</summary>
    public List<CellBorderOptions>? Borders { get; set; }
    /// <summary>left, center, right, justify, distributed, fill, or centerAcrossSelection.</summary>
    public string? HorizontalAlignment { get; set; }
    /// <summary>top, center, middle, bottom, justify, or distributed.</summary>
    public string? VerticalAlignment { get; set; }
    /// <summary>Text wrapping.</summary>
    public bool? WrapText { get; set; }
    /// <summary>Text shrinking to fit.</summary>
    public bool? ShrinkToFit { get; set; }
    /// <summary>Native indentation level, 0 through 15.</summary>
    public int? IndentLevel { get; set; }
    /// <summary>context, leftToRight, or rightToLeft.</summary>
    public string? ReadingOrder { get; set; }
    /// <summary>Text rotation -90 through 90, or 255 for vertical text.</summary>
    public int? Orientation { get; set; }
    /// <summary>Invariant Excel number-format code, translated for the installed Excel locale.</summary>
    public string? NumberFormat { get; set; }
}
