using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Formatting operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    private static readonly int[] BorderEdges = [7, 8, 9, 10];

    /// <inheritdoc />
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2012:Use ValueTasks correctly")]
    public void SetStyle(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string styleName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Apply built-in style
                range.Style = styleName;

                return ValueTask.CompletedTask;
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public RangeStyleResult GetStyle(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get style name from the first cell in the range
                string styleName;
                try
                {
                    styleName = ComUtilities.SafeGetString(range.Style, "Name");
                    if (string.IsNullOrEmpty(styleName))
                        styleName = "Normal";
                }
                catch
                {
                    styleName = "Normal";
                }

                // Try to determine if it's a built-in style
                bool isBuiltIn = false;
                string? styleDescription = null;

                try
                {
                    dynamic styles = ctx.Book.Styles;
                    dynamic style = styles.Item(styleName);
                    isBuiltIn = true;

                    // Try to get additional information about the style
                    try
                    {
                        styleDescription = ComUtilities.SafeGetString(style, "NameLocal");
                        if (string.IsNullOrEmpty(styleDescription))
                            styleDescription = null;
                    }
                    catch
                    {
                        // Style description is optional
                    }
                }
                catch
                {
                    // If we can't find it in the Styles collection, it might be a custom style
                    isBuiltIn = false;
                }

                return new RangeStyleResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    RangeAddress = range.Address,
                    StyleName = styleName,
                    IsBuiltInStyle = isBuiltIn,
                    StyleDescription = styleDescription
                };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2012:Use ValueTasks correctly")]
    public void FormatRange(
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
        batch.Execute((ctx, ct) =>
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
                    if (fontColor != null) font.Color = FormattingHelpers.ParseColor(fontColor);
                }

                // Apply fill color
                if (fillColor != null)
                {
                    interior = range.Interior;
                    interior.Color = FormattingHelpers.ParseColor(fillColor);
                }

                // Apply borders
                if (borderStyle != null || borderColor != null || borderWeight != null)
                {
                    // Apply to all edges (7 = xlEdgeLeft, 8 = xlEdgeTop, 9 = xlEdgeBottom, 10 = xlEdgeRight)
                    foreach (var edge in BorderEdges)
                    {
                        dynamic? border = null;
                        try
                        {
                            border = range.Borders.Item(edge);
                            if (borderStyle != null) border.LineStyle = FormattingHelpers.ParseBorderStyle(borderStyle);
                            if (borderColor != null) border.Color = FormattingHelpers.ParseColor(borderColor);
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

                return ValueTask.CompletedTask;
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



