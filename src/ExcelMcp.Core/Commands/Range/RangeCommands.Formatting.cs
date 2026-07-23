using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Formatting operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    private static readonly int[] BorderEdges = [7, 8, 9, 10];

    private readonly record struct RangeFormatRequest(
        string? FontName,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        int? FontColor,
        int? FillColor,
        int? BorderStyle,
        int? BorderColor,
        int? BorderWeight,
        int? HorizontalAlignment,
        int? VerticalAlignment,
        bool? WrapText,
        int? Orientation,
        string? NumberFormat)
    {
        public bool HasFontFormatting => FontName != null || FontSize != null || Bold != null || Italic != null || Underline != null || FontColor != null;

        public bool HasBorderFormatting => BorderStyle != null || BorderColor != null || BorderWeight != null;
    }

    /// <inheritdoc />
    public OperationResult SetStyle(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string styleName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                range = sheet.Range[rangeAddress];
                range.Style = styleName;

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "set-style"
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
    public RangeStyleResult GetStyle(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? styles = null;
            dynamic? style = null;

            try
            {
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                range = sheet.Range[rangeAddress];

                string styleName;
                try
                {
                    styleName = ComUtilities.SafeGetString(range.Style, "Name");
                    if (string.IsNullOrEmpty(styleName))
                    {
                        styleName = "Normal";
                    }
                }
                catch (COMException)
                {
                    styleName = "Normal";
                }

                var isBuiltIn = false;
                string? styleDescription = null;

                try
                {
                    styles = ctx.Book.Styles;
                    style = styles.Item(styleName);
                    isBuiltIn = true;

                    try
                    {
                        styleDescription = ComUtilities.SafeGetString(style, "NameLocal");
                        if (string.IsNullOrEmpty(styleDescription))
                        {
                            styleDescription = null;
                        }
                    }
                    catch (COMException)
                    {
                    }
                }
                catch (COMException)
                {
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
                ComUtilities.Release(ref style!);
                ComUtilities.Release(ref styles!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult FormatRange(
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
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                var formatRequest = CreateFormatRequest(
                    fontName,
                    fontSize,
                    bold,
                    italic,
                    underline,
                    fontColor,
                    fillColor,
                    borderStyle,
                    borderColor,
                    borderWeight,
                    horizontalAlignment,
                    verticalAlignment,
                    wrapText,
                    orientation);

                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                range = sheet.Range[rangeAddress];
                ApplyFormattingToRange(range, formatRequest);

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "format-range"
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
    public OperationResult FormatRanges(
        IExcelBatch batch,
        string sheetName,
        string[] rangeAddresses,
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
        int? orientation,
        string? numberFormat = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;

            try
            {
                ValidateRangeAddresses(rangeAddresses);

                var formatRequest = CreateFormatRequest(
                    fontName,
                    fontSize,
                    bold,
                    italic,
                    underline,
                    fontColor,
                    fillColor,
                    borderStyle,
                    borderColor,
                    borderWeight,
                    horizontalAlignment,
                    verticalAlignment,
                    wrapText,
                    orientation,
                    numberFormat);

                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                ValidateTargetRanges(sheet, rangeAddresses, nameof(rangeAddresses));

                foreach (var rangeAddress in rangeAddresses)
                {
                    dynamic? range = null;

                    try
                    {
                        range = sheet.Range[rangeAddress];
                        ApplyFormattingToRange(range, formatRequest);
                    }
                    finally
                    {
                        ComUtilities.Release(ref range!);
                    }
                }

                return new OperationResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    Action = "format-ranges"
                };
            }
            finally
            {
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private static RangeFormatRequest CreateFormatRequest(
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
        int? orientation,
        string? numberFormat = null)
    {
        return new RangeFormatRequest(
            fontName,
            fontSize,
            bold,
            italic,
            underline,
            fontColor is null ? null : FormattingHelpers.ParseColor(fontColor),
            fillColor is null ? null : FormattingHelpers.ParseColor(fillColor),
            borderStyle is null ? null : FormattingHelpers.ParseBorderStyle(borderStyle),
            borderColor is null ? null : FormattingHelpers.ParseColor(borderColor),
            borderWeight is null ? null : ParseBorderWeight(borderWeight),
            horizontalAlignment is null ? null : ParseHorizontalAlignment(horizontalAlignment),
            verticalAlignment is null ? null : ParseVerticalAlignment(verticalAlignment),
            wrapText,
            orientation,
            numberFormat);
    }

    private static void ApplyFormattingToRange(dynamic range, RangeFormatRequest formatRequest)
    {
        dynamic? font = null;
        dynamic? interior = null;

        try
        {
            if (formatRequest.HasFontFormatting)
            {
                font = range.Font;
                if (formatRequest.FontName != null) font.Name = formatRequest.FontName;
                if (formatRequest.FontSize != null) font.Size = formatRequest.FontSize.Value;
                if (formatRequest.Bold != null) font.Bold = formatRequest.Bold.Value;
                if (formatRequest.Italic != null) font.Italic = formatRequest.Italic.Value;
                if (formatRequest.Underline != null) font.Underline = formatRequest.Underline.Value ? 2 : -4142;
                if (formatRequest.FontColor != null) font.Color = formatRequest.FontColor.Value;
            }

            if (formatRequest.FillColor != null)
            {
                interior = range.Interior;
                interior.Color = formatRequest.FillColor.Value;
            }

            if (formatRequest.HasBorderFormatting)
            {
                foreach (var edge in BorderEdges)
                {
                    dynamic? border = null;

                    try
                    {
                        border = range.Borders.Item(edge);
                        if (formatRequest.BorderStyle != null) border.LineStyle = formatRequest.BorderStyle.Value;
                        if (formatRequest.BorderColor != null) border.Color = formatRequest.BorderColor.Value;
                        if (formatRequest.BorderWeight != null) border.Weight = formatRequest.BorderWeight.Value;
                    }
                    finally
                    {
                        ComUtilities.Release(ref border!);
                    }
                }
            }

            if (formatRequest.HorizontalAlignment != null)
            {
                range.HorizontalAlignment = formatRequest.HorizontalAlignment.Value;
            }

            if (formatRequest.VerticalAlignment != null)
            {
                range.VerticalAlignment = formatRequest.VerticalAlignment.Value;
            }

            if (formatRequest.WrapText != null)
            {
                range.WrapText = formatRequest.WrapText.Value;
            }

            if (formatRequest.Orientation != null)
            {
                range.Orientation = formatRequest.Orientation.Value;
            }

            if (formatRequest.NumberFormat != null)
            {
                range.NumberFormat = formatRequest.NumberFormat;
            }
        }
        finally
        {
            ComUtilities.Release(ref interior!);
            ComUtilities.Release(ref font!);
        }
    }

    private static void ValidateRangeAddresses(string[] rangeAddresses)
    {
        if (rangeAddresses == null || rangeAddresses.Length == 0)
        {
            throw new ArgumentException("At least one range address is required.", nameof(rangeAddresses));
        }

        for (var index = 0; index < rangeAddresses.Length; index++)
        {
            if (string.IsNullOrWhiteSpace(rangeAddresses[index]))
            {
                throw new ArgumentException($"Range address at index {index} cannot be null, empty, or whitespace.", nameof(rangeAddresses));
            }
        }
    }

    private static void ValidateTargetRanges(dynamic sheet, IReadOnlyList<string> rangeAddresses, string parameterName)
    {
        for (var index = 0; index < rangeAddresses.Count; index++)
        {
            dynamic? range = null;

            try
            {
                range = sheet.Range[rangeAddresses[index]];
                _ = range.Address;
            }
            catch (COMException ex)
            {
                throw new ArgumentException($"Invalid range address at index {index}: '{rangeAddresses[index]}'", parameterName, ex);
            }
            finally
            {
                ComUtilities.Release(ref range!);
            }
        }
    }

    private static int ParseBorderWeight(string weight)
    {
        return weight.ToLowerInvariant() switch
        {
            "hairline" => 1,
            "thin" => 2,
            "medium" => -4138,
            "thick" => 4,
            _ => throw new ArgumentException($"Invalid border weight: {weight}")
        };
    }

    private static int ParseHorizontalAlignment(string alignment)
    {
        return alignment.ToLowerInvariant() switch
        {
            "left" => -4131,
            "center" => -4108,
            "right" => -4152,
            "justify" => -4130,
            "distributed" => -4117,
            _ => throw new ArgumentException($"Invalid horizontal alignment: {alignment}")
        };
    }

    private static int ParseVerticalAlignment(string alignment)
    {
        return alignment.ToLowerInvariant() switch
        {
            "top" => -4160,
            "center" => -4108,
            "middle" => -4108,
            "bottom" => -4107,
            "justify" => -4130,
            "distributed" => -4117,
            _ => throw new ArgumentException($"Invalid vertical alignment: {alignment}")
        };
    }
}



