using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Formatting operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> FormatRangeAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string? fontName,
        double? fontSize,
        bool? bold,
        bool? italic,
        bool? underline,
        string? fontColor,
        string? fillColor,
        string? borderStyle,
        string? borderColor,
        string? borderWeight,
        string? horizontalAlignment,
        string? verticalAlignment,
        bool? wrapText,
        int? orientation)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            dynamic? interior = null;
            dynamic? borders = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Apply font formatting
                if (fontName != null || fontSize != null || bold != null || italic != null || underline != null || fontColor != null)
                {
                    font = range.Font;
                    if (fontName != null) font.Name = fontName;
                    if (fontSize != null) font.Size = fontSize.Value;
                    if (bold != null) font.Bold = bold.Value;
                    if (italic != null) font.Italic = italic.Value;
                    if (underline != null) font.Underline = underline.Value ? 2 : -4142; // xlUnderlineStyleSingle : xlUnderlineStyleNone
                    if (fontColor != null) font.Color = ParseColor(fontColor);
                }

                // Apply fill color
                if (fillColor != null)
                {
                    interior = range.Interior;
                    interior.Color = ParseColor(fillColor);
                }

                // Apply borders
                if (borderStyle != null || borderColor != null || borderWeight != null)
                {
                    // Apply to all edges (7 = xlEdgeLeft, 8 = xlEdgeTop, 9 = xlEdgeBottom, 10 = xlEdgeRight)
                    int[] edges = { 7, 8, 9, 10 };
                    foreach (var edge in edges)
                    {
                        dynamic? border = null;
                        try
                        {
                            border = range.Borders.Item(edge);
                            if (borderStyle != null) border.LineStyle = ParseBorderStyle(borderStyle);
                            if (borderColor != null) border.Color = ParseColor(borderColor);
                            if (borderWeight != null) border.Weight = ParseBorderWeight(borderWeight);
                        }
                        finally
                        {
                            ComUtilities.Release(ref border!);
                        }
                    }
                }

                // Apply alignment
                if (horizontalAlignment != null)
                {
                    range.HorizontalAlignment = ParseHorizontalAlignment(horizontalAlignment);
                }
                if (verticalAlignment != null)
                {
                    range.VerticalAlignment = ParseVerticalAlignment(verticalAlignment);
                }

                // Apply text wrapping
                if (wrapText != null)
                {
                    range.WrapText = wrapText.Value;
                }

                // Apply orientation
                if (orientation != null)
                {
                    range.Orientation = orientation.Value;
                }

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to format range '{rangeAddress}': {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref borders!);
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private static int ParseColor(string color)
    {
        // Support #RRGGBB format or color index
        if (color.StartsWith("#") && color.Length == 7)
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

    private static int ParseBorderStyle(string style)
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

    private static int ParseBorderWeight(string weight)
    {
        return weight.ToLowerInvariant() switch
        {
            "hairline" => 1, // xlHairline
            "thin" => 2, // xlThin
            "medium" => -4138, // xlMedium
            "thick" => 4, // xlThick
            _ => throw new ArgumentException($"Invalid border weight: {weight}")
        };
    }

    private static int ParseHorizontalAlignment(string alignment)
    {
        return alignment.ToLowerInvariant() switch
        {
            "left" => -4131, // xlLeft
            "center" => -4108, // xlCenter
            "right" => -4152, // xlRight
            "justify" => -4130, // xlJustify
            "distributed" => -4117, // xlDistributed
            _ => throw new ArgumentException($"Invalid horizontal alignment: {alignment}")
        };
    }

    private static int ParseVerticalAlignment(string alignment)
    {
        return alignment.ToLowerInvariant() switch
        {
            "top" => -4160, // xlTop
            "center" => -4108, // xlCenter
            "bottom" => -4107, // xlBottom
            "justify" => -4130, // xlJustify
            "distributed" => -4117, // xlDistributed
            _ => throw new ArgumentException($"Invalid vertical alignment: {alignment}")
        };
    }
}
