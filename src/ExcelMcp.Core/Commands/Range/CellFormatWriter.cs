using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

internal static class CellFormatWriter
{
    internal static void Validate(CellFormatOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        if (options.FontName is not null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(options.FontName);
            if (options.ThemeFont is not null)
                throw new ArgumentException("fontName and themeFont are mutually exclusive.");
        }
        if (options.FontSize is { } size && (!double.IsFinite(size) || size < 1 || size > 409))
            throw new ArgumentException("fontSize must be finite and between 1 and 409.");
        if (options.ThemeFont is < 0 or > 2)
            throw new ArgumentException("themeFont must be 0, 1, or 2.");
        if (options.Subscript == true && options.Superscript == true)
            throw new ArgumentException("subscript and superscript cannot both be true.");
        if (options.IndentLevel is < 0 or > 15)
            throw new ArgumentException("indentLevel must be between 0 and 15.");
        if (options.Orientation is { } orientation && orientation != 255 && (orientation < -90 || orientation > 90))
            throw new ArgumentException("orientation must be -90 through 90, or 255.");
        ValidateColor(options.FontColor, options.FontThemeColor, options.FontTintAndShade, "font");
        ValidateColor(options.FillColor, options.FillThemeColor, options.FillTintAndShade, "fill");
        if (options.Underline is not null) _ = ParseUnderline(options.Underline);
        if (options.HorizontalAlignment is not null) _ = ParseHorizontalAlignment(options.HorizontalAlignment);
        if (options.VerticalAlignment is not null) _ = ParseVerticalAlignment(options.VerticalAlignment);
        if (options.ReadingOrder is not null) _ = ParseReadingOrder(options.ReadingOrder);
        if (options.NumberFormat is not null) ArgumentException.ThrowIfNullOrWhiteSpace(options.NumberFormat);
        HashSet<CellBorderPosition> positions = [];
        foreach (var border in options.Borders ?? [])
        {
            if (border is null || !Enum.IsDefined(border.Position) || !positions.Add(border.Position))
                throw new ArgumentException("Each border must specify a valid, unique position.");
            if (border.LineStyle is not null) _ = FormattingHelpers.ParseBorderStyle(border.LineStyle);
            if (border.Weight is not null) _ = ParseBorderWeight(border.Weight);
            if (border.LineStyle is null && border.Weight is null && border.Color is null &&
                border.ThemeColor is null && border.TintAndShade is null)
                throw new ArgumentException("Each border must specify at least one setting.");
            if (border.LineStyle?.Equals("none", StringComparison.OrdinalIgnoreCase) == true &&
                (border.Weight is not null || border.Color is not null || border.ThemeColor is not null || border.TintAndShade is not null))
                throw new ArgumentException("A border with lineStyle none cannot also specify weight or color.");
            ValidateColor(border.Color, border.ThemeColor, border.TintAndShade, $"border {border.Position}");
        }
    }

    private static void ValidateColor(string? rgb, int? theme, double? tint, string field)
    {
        if (rgb is not null) _ = FormattingHelpers.ParseColor(rgb);
        if (rgb is not null && theme is not null)
            throw new ArgumentException($"{field}: fixed RGB and theme color are mutually exclusive.");
        if (theme is < 1 or > 12)
            throw new ArgumentException($"{field}: theme color must be between 1 and 12.");
        if (tint is { } value && (!double.IsFinite(value) || value < -1 || value > 1))
            throw new ArgumentException($"{field}: tint/shade must be finite and between -1 and 1.");
    }

    internal static void ApplyFont(Excel.Font font, CellFormatOptions options)
    {
        if (options.FontName is not null) font.Name = options.FontName;
        if (options.ThemeFont is { } themeFont) font.ThemeFont = (Excel.XlThemeFont)themeFont;
        if (options.FontSize is { } size) font.Size = size;
        if (options.Bold is { } bold) font.Bold = bold;
        if (options.Italic is { } italic) font.Italic = italic;
        if (options.Underline is not null) font.Underline = ParseUnderline(options.Underline);
        if (options.Strikethrough is { } strike) font.Strikethrough = strike;
        if (options.Subscript is { } subscript) font.Subscript = subscript;
        if (options.Superscript is { } superscript) font.Superscript = superscript;
        if (options.FontColor is not null) font.Color = FormattingHelpers.ParseColor(options.FontColor);
        if (options.FontThemeColor is { } theme) font.ThemeColor = theme;
        if (options.FontTintAndShade is { } tint) font.TintAndShade = tint;
    }

    internal static void ApplyFill(Excel.Interior fill, CellFormatOptions options)
    {
        if (options.FillColor is not null) fill.Color = FormattingHelpers.ParseColor(options.FillColor);
        if (options.FillThemeColor is { } theme) fill.ThemeColor = theme;
        if (options.FillTintAndShade is { } tint) fill.TintAndShade = tint;
    }

    internal static void ApplyBorders(Excel.Borders borders, CellFormatOptions options, CancellationToken ct,
        bool styleDefinition = false)
    {
        foreach (var setting in options.Borders ?? [])
        {
            ct.ThrowIfCancellationRequested();
            Excel.Border? border = null;
            try
            {
                border = borders[styleDefinition ? StyleBorderIndex(setting.Position) : (Excel.XlBordersIndex)setting.Position];
                // Set weight before style: Excel can otherwise replace a requested double/dashed line.
                if (setting.Weight is not null) border.Weight = ParseBorderWeight(setting.Weight);
                if (setting.LineStyle is not null) border.LineStyle = FormattingHelpers.ParseBorderStyle(setting.LineStyle);
                // Enabling a previously absent style border can reset its color to automatic.
                if (setting.Color is not null) border.Color = FormattingHelpers.ParseColor(setting.Color);
                if (setting.ThemeColor is { } theme) border.ThemeColor = theme;
                if (setting.TintAndShade is { } tint) border.TintAndShade = tint;
            }
            finally
            {
                ComUtilities.Release(ref border);
            }
        }
    }

    internal static Excel.XlBordersIndex StyleBorderIndex(CellBorderPosition position) => position switch
    {
        // Style.Borders uses legacy side constants, unlike Range.Borders' xlEdge* indices.
        CellBorderPosition.Left => (Excel.XlBordersIndex)(-4131),
        CellBorderPosition.Top => (Excel.XlBordersIndex)(-4160),
        CellBorderPosition.Bottom => (Excel.XlBordersIndex)(-4107),
        CellBorderPosition.Right => (Excel.XlBordersIndex)(-4152),
        CellBorderPosition.DiagonalDown or CellBorderPosition.DiagonalUp => (Excel.XlBordersIndex)position,
        _ => throw new ArgumentException("Cell styles do not define inside borders. Use range_format format on the requested range.")
    };

    internal static int ParseUnderline(string value) => value.ToLowerInvariant() switch
    {
        "none" => -4142,
        "single" => 2,
        "double" => -4119,
        "singleaccounting" => 4,
        "doubleaccounting" => 5,
        _ => throw new ArgumentException($"Invalid underline: {value}")
    };

    internal static int ParseBorderWeight(string value) => value.ToLowerInvariant() switch
    {
        "hairline" => 1,
        "thin" => 2,
        "medium" => -4138,
        "thick" => 4,
        _ => throw new ArgumentException($"Invalid border weight: {value}")
    };

    internal static int ParseHorizontalAlignment(string value) => value.ToLowerInvariant() switch
    {
        "left" => -4131,
        "center" => -4108,
        "right" => -4152,
        "justify" => -4130,
        "distributed" => -4117,
        "fill" => 5,
        "centeracrossselection" => 7,
        _ => throw new ArgumentException($"Invalid horizontal alignment: {value}")
    };

    internal static int ParseVerticalAlignment(string value) => value.ToLowerInvariant() switch
    {
        "top" => -4160,
        "center" or "middle" => -4108,
        "bottom" => -4107,
        "justify" => -4130,
        "distributed" => -4117,
        _ => throw new ArgumentException($"Invalid vertical alignment: {value}")
    };

    internal static int ParseReadingOrder(string value) => value.ToLowerInvariant() switch
    {
        "context" => -5002,
        "lefttoright" => -5003,
        "righttoleft" => -5004,
        _ => throw new ArgumentException($"Invalid reading order: {value}")
    };
}
