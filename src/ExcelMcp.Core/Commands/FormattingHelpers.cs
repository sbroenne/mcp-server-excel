namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Shared formatting helper methods for Excel formatting operations.
/// Used by Range and ConditionalFormat commands.
/// </summary>
internal static class FormattingHelpers
{
    /// <summary>
    /// Parses a color string to Excel RGB format.
    /// Supports #RRGGBB format or color index.
    /// </summary>
    /// <param name="color">Color in #RRGGBB format or numeric index</param>
    /// <returns>Excel RGB integer value</returns>
    /// <exception cref="ArgumentException">If color format is invalid</exception>
    public static int ParseColor(string color)
    {
        // Support #RRGGBB format or color index
        if (color.StartsWith('#') && color.Length == 7)
        {
            var r = Convert.ToInt32(color.Substring(1, 2), 16);
            var g = Convert.ToInt32(color.Substring(3, 2), 16);
            var b = Convert.ToInt32(color.Substring(5, 2), 16);
            return r + (g << 8) + (b << 16); // Excel RGB format
        }
        else if (int.TryParse(color, out var index))
        {
            return index;
        }
        throw new ArgumentException($"Invalid color format: {color}. Use #RRGGBB or color index.");
    }

    /// <summary>
    /// Parses a border style string to Excel constant.
    /// </summary>
    /// <param name="style">Border style name</param>
    /// <returns>Excel border style constant</returns>
    /// <exception cref="ArgumentException">If style is invalid</exception>
    public static int ParseBorderStyle(string style)
    {
        return style.ToLowerInvariant() switch
        {
            "none" => -4142, // xlNone
            "continuous" => 1, // xlContinuous
            "dash" => -4115, // xlDash
            "dashdot" => 4, // xlDashDot
            "dashdotdot" => 5, // xlDashDotDot
            "dot" => -4118, // xlDot
            "double" => -4119, // xlDouble
            "slantdashdot" => 13, // xlSlantDashDot
            _ => throw new ArgumentException($"Invalid border style: {style}")
        };
    }
}
