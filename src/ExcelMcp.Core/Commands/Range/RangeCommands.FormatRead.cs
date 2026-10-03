using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    private static readonly Excel.XlBordersIndex[] FormatReadBorders =
    [
        Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeTop,
        Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeRight,
        Excel.XlBordersIndex.xlDiagonalDown, Excel.XlBordersIndex.xlDiagonalUp,
        Excel.XlBordersIndex.xlInsideHorizontal, Excel.XlBordersIndex.xlInsideVertical
    ];

    /// <inheritdoc />
    public RangeFormatReadResult GetFormat(
        IExcelBatch batch, string sheetName, string rangeAddress, FormatView view = FormatView.Stored)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        if (!Enum.IsDefined(view))
        {
            throw new ArgumentOutOfRangeException(nameof(view));
        }
        return batch.Execute((ctx, ct) =>
        {
            Excel.Range? range = null;
            Excel.Worksheet? sheet = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                sheet = (Excel.Worksheet)range!.Parent;
                var result = new RangeFormatReadResult
                {
                    FilePath = batch.WorkbookPath,
                    SheetName = sheet.Name,
                    RangeAddress = range.Address,
                    View = view
                };
                RangeHelpers.VisitCells(range, ct, cell =>
                    result.Cells.Add(new CellFormatRead(cell.Address, cell.Row, cell.Column,
                        view is FormatView.Stored or FormatView.Both ? ReadCellFormat(cell, false, ct) : null,
                        view is FormatView.Displayed or FormatView.Both ? ReadCellFormat(cell, true, ct) : null)));
                result.Cells = result.Cells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column).ToList();
                result.CellCount = result.Cells.Count;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref range);
            }
        });
    }

    private static CellFormatSnapshot ReadCellFormat(Excel.Range cell, bool displayed, CancellationToken ct)
    {
        Excel.DisplayFormat? display = null;
        Excel.Font? font = null;
        Excel.Interior? interior = null;
        Excel.Borders? borders = null;
        try
        {
            if (displayed)
            {
                display = cell.DisplayFormat;
            }
            font = display is null ? cell.Font : display.Font;
            interior = display is null ? cell.Interior : display.Interior;
            borders = display is null ? cell.Borders : display.Borders;
            List<string> mixed = [];
            var (fontFormat, fillFormat, borderFormats) = ReadVisualFormat(font, interior, borders, mixed, ct);
            return new CellFormatSnapshot
            {
                Font = fontFormat,
                Fill = fillFormat,
                Borders = borderFormats,
                NumberFormat = ReadFormatText(display is null ? (object)cell.NumberFormat : (object)display.NumberFormat, "numberFormat", mixed),
                HorizontalAlignment = ReadFormatValue<int>(display is null ? (object)cell.HorizontalAlignment : (object)display.HorizontalAlignment, "horizontalAlignment", mixed),
                VerticalAlignment = ReadFormatValue<int>(display is null ? (object)cell.VerticalAlignment : (object)display.VerticalAlignment, "verticalAlignment", mixed),
                WrapText = ReadFormatValue<bool>(display is null ? (object)cell.WrapText : (object)display.WrapText, "wrapText", mixed),
                ShrinkToFit = ReadFormatValue<bool>(display is null ? (object)cell.ShrinkToFit : (object)display.ShrinkToFit, "shrinkToFit", mixed),
                AddIndent = ReadFormatValue<bool>(display is null ? (object)cell.AddIndent : (object)display.AddIndent, "addIndent", mixed),
                IndentLevel = ReadFormatValue<int>(display is null ? (object)cell.IndentLevel : (object)display.IndentLevel, "indentLevel", mixed),
                Orientation = ReadFormatValue<int>(display is null ? (object)cell.Orientation : (object)display.Orientation, "orientation", mixed),
                ReadingOrder = ReadFormatValue<int>(display is null ? (object)cell.ReadingOrder : (object)display.ReadingOrder, "readingOrder", mixed),
                Locked = ReadFormatValue<bool>(display is null ? (object)cell.Locked : (object)display.Locked, "locked", mixed),
                FormulaHidden = ReadFormatValue<bool>(display is null ? (object)cell.FormulaHidden : (object)display.FormulaHidden, "formulaHidden", mixed),
                MergeCells = ReadFormatValue<bool>(display is null ? (object)cell.MergeCells : (object)display.MergeCells, "mergeCells", mixed),
                StyleName = ReadCellStyle(cell, display, mixed),
                ColumnWidth = ReadFormatValue<double>((object)cell.ColumnWidth, "columnWidth", mixed),
                RowHeight = ReadFormatValue<double>((object)cell.RowHeight, "rowHeight", mixed),
                MixedFields = mixed
            };
        }
        finally
        {
            ComUtilities.Release(ref borders);
            ComUtilities.Release(ref interior);
            ComUtilities.Release(ref font);
            ComUtilities.Release(ref display);
        }
    }


    internal static (CellFontFormat Font, CellFillFormat Fill, List<CellBorderFormat> Borders) ReadVisualFormat(
        Excel.Font font, Excel.Interior interior, Excel.Borders borders, List<string> mixed, CancellationToken ct,
        bool styleDefinition = false, bool tableStyleDefinition = false)
    {
        object? Read(Func<object?> getter, string field)
        {
            if (!styleDefinition && !tableStyleDefinition) return getter();
            try { return getter(); }
            catch (Exception ex) when (ex is COMException or ArgumentException &&
                ex.HResult is unchecked((int)0x800A03EC) or unchecked((int)0x80070057))
            {
                mixed.Add($"{field}.unavailable");
                return null;
            }
        }
        object? ReadThemeFont()
        {
            if (!tableStyleDefinition) return font.ThemeFont;
            // PIA gap: differential Font.ThemeFont can be native null, unlike the declared value-type enum.
            return ((dynamic)font).ThemeFont;
        }
        var fontFormat = new CellFontFormat(
            ReadFormatText(Read(() => font.Name, "font.name"), "font.name", mixed),
            ReadFormatValue<double>(Read(() => font.Size, "font.size"), "font.size", mixed),
            ReadFormatValue<bool>((object)font.Bold, "font.bold", mixed),
            ReadFormatValue<bool>((object)font.Italic, "font.italic", mixed),
            ReadFormatValue<int>((object)font.Underline, "font.underline", mixed),
            ReadFormatValue<bool>((object)font.Strikethrough, "font.strikethrough", mixed),
            ReadFormatValue<bool>(Read(() => font.Subscript, "font.subscript"), "font.subscript", mixed),
            ReadFormatValue<bool>(Read(() => font.Superscript, "font.superscript"), "font.superscript", mixed),
            ReadFormatValue<int>(Read(ReadThemeFont, "font.themeFont"), "font.themeFont", mixed),
            ReadFormatColor(Read(() => font.Color, "font.color.rgb"), Read(() => font.ColorIndex, "font.color.colorIndex"), Read(() => font.TintAndShade, "font.color.tintAndShade"),
                () => font.ThemeColor, "font.color", mixed));
        var fillFormat = new CellFillFormat(
            ReadFormatValue<int>((object)interior.Pattern, "fill.pattern", mixed),
            ReadFormatColor(Read(() => interior.Color, "fill.color.rgb"), Read(() => interior.ColorIndex, "fill.color.colorIndex"), Read(() => interior.TintAndShade, "fill.color.tintAndShade"),
                () => interior.ThemeColor, "fill.color", mixed),
            ReadFormatColor(Read(() => interior.PatternColor, "fill.patternColor.rgb"), Read(() => interior.PatternColorIndex, "fill.patternColor.colorIndex"), Read(() => interior.PatternTintAndShade, "fill.patternColor.tintAndShade"),
                () => interior.PatternThemeColor, "fill.patternColor", mixed),
            ReadCellGradient(interior, mixed, ct));
        List<CellBorderFormat> borderFormats = [];
        foreach (var edge in FormatReadBorders)
        {
            if (styleDefinition && edge is Excel.XlBordersIndex.xlInsideHorizontal or Excel.XlBordersIndex.xlInsideVertical)
                continue;
            if (tableStyleDefinition && edge is Excel.XlBordersIndex.xlDiagonalDown or Excel.XlBordersIndex.xlDiagonalUp)
                continue;
            ct.ThrowIfCancellationRequested();
            Excel.Border? border = null;
            try
            {
                border = borders[styleDefinition ? CellFormatWriter.StyleBorderIndex((CellBorderPosition)edge) : edge];
                string field = $"borders.{edge}";
                borderFormats.Add(new CellBorderFormat(edge.ToString(),
                    ReadFormatValue<int>((object)border.LineStyle, $"{field}.lineStyle", mixed),
                    ReadFormatValue<int>((object)border.Weight, $"{field}.weight", mixed),
                    ReadFormatColor(Read(() => border.Color, $"{field}.color.rgb"), Read(() => border.ColorIndex, $"{field}.color.colorIndex"), Read(() => border.TintAndShade, $"{field}.color.tintAndShade"),
                        () => border.ThemeColor, $"{field}.color", mixed)));
            }
            finally
            {
                ComUtilities.Release(ref border);
            }
        }
        return (fontFormat, fillFormat, borderFormats);
    }

    private static string? ReadCellStyle(
        Excel.Range cell, Excel.DisplayFormat? display, List<string> mixed)
    {
        object? style = display is null ? cell.Style : display.Style;
        try
        {
            return style is Excel.Style nativeStyle
                ? nativeStyle.Name
                : ReadFormatText(style, "styleName", mixed);
        }
        finally
        {
            if (style is not null && Marshal.IsComObject(style))
            {
                ComUtilities.Release(ref style);
            }
        }
    }

    private static CellGradientFormat? ReadCellGradient(
        Excel.Interior interior, List<string> mixed, CancellationToken ct)
    {
        object? patternValue = interior.Pattern;
        if (patternValue is null or DBNull) return null;
        int pattern = Convert.ToInt32(patternValue, CultureInfo.InvariantCulture);
        if (pattern is not ((int)Excel.XlPattern.xlPatternLinearGradient)
            and not ((int)Excel.XlPattern.xlPatternRectangularGradient))
        {
            return null;
        }
        object? gradient = null;
        Excel.ColorStops? stops = null;
        try
        {
            gradient = interior.Gradient;
            string kind;
            double? degree = null, top = null, bottom = null, left = null, right = null;
            if (pattern == (int)Excel.XlPattern.xlPatternLinearGradient)
            {
                var linear = (Excel.LinearGradient)gradient;
                kind = "linear";
                degree = linear.Degree;
                stops = linear.ColorStops;
            }
            else
            {
                var rectangular = (Excel.RectangularGradient)gradient;
                kind = "rectangular";
                top = rectangular.RectangleTop;
                bottom = rectangular.RectangleBottom;
                left = rectangular.RectangleLeft;
                right = rectangular.RectangleRight;
                stops = rectangular.ColorStops;
            }
            List<CellGradientStop> result = [];
            for (int index = 1; index <= stops.Count; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.ColorStop? stop = null;
                try
                {
                    stop = stops[index];
                    int color = Convert.ToInt32(stop.Color, CultureInfo.InvariantCulture);
                    result.Add(new CellGradientStop(stop.Position,
                        new CellColorFormat(FormattingHelpers.ColorToHex(color), null,
                            ReadFormatTheme(() => stop.ThemeColor, $"fill.gradient.stops[{index}].themeColor", mixed),
                            stop.TintAndShade)));
                }
                finally
                {
                    ComUtilities.Release(ref stop);
                }
            }
            return new CellGradientFormat(kind, degree, top, bottom, left, right,
                result.OrderBy(stop => stop.Position).ToList());
        }
        finally
        {
            ComUtilities.Release(ref stops);
            ComUtilities.Release(ref gradient);
        }
    }

    private static T? ReadFormatValue<T>(object? value, string field, List<string> mixed) where T : struct
    {
        if (value is null or DBNull)
        {
            mixed.Add(field);
            return null;
        }
        return (T)Convert.ChangeType(value, typeof(T), CultureInfo.InvariantCulture);
    }

    private static string? ReadFormatText(object? value, string field, List<string> mixed)
    {
        if (value is null or DBNull)
        {
            mixed.Add(field);
            return null;
        }
        return Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static CellColorFormat ReadFormatColor(
        object? color, object? colorIndex, object? tint, Func<object> theme,
        string field, List<string> mixed)
    {
        var index = ReadFormatValue<int>(colorIndex, $"{field}.colorIndex", mixed);
        if (index == 0)
        {
            mixed.Add(field);
            return new CellColorFormat(null, index, null,
                ReadFormatValue<double>(tint, $"{field}.tintAndShade", mixed));
        }
        int? themeIndex = null;
        if (index.HasValue && index != (int)Excel.XlColorIndex.xlColorIndexNone)
        {
            themeIndex = ReadFormatTheme(theme, $"{field}.themeColor", mixed);
        }
        var nativeColor = ReadFormatValue<int>(color, $"{field}.rgb", mixed);
        return new CellColorFormat(
            index is null or (int)Excel.XlColorIndex.xlColorIndexNone || nativeColor is null
                ? null : FormattingHelpers.ColorToHex(nativeColor.Value),
            index, themeIndex,
            ReadFormatValue<double>(tint, $"{field}.tintAndShade", mixed));
    }

    private static int? ReadFormatTheme(Func<object> theme, string field, List<string> mixed)
    {
        try
        {
            return ReadFormatValue<int>(theme(), field, mixed);
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
        {
            // Excel's ThemeColor getter rejects valid colors that are not theme-based.
            return null;
        }
        catch (ArgumentException ex) when (ex.HResult == unchecked((int)0x80070057))
        {
            // The PIA also maps the native non-theme-color rejection to E_INVALIDARG.
            return null;
        }
    }
}
