using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

public partial class ChartCommands
{
    /// <inheritdoc />
    public ChartTextFormatResult GetTextFormat(IExcelBatch batch, string chartName,
        ChartTextElement element, ChartAxisType? axis = null, int? seriesIndex = null)
    {
        axis = ValidateTextTarget(chartName, element, axis, seriesIndex);
        return WithTextFormats(batch, chartName, element, axis, seriesIndex, null);
    }

    /// <inheritdoc />
    public ChartTextFormatResult SetTextFormat(IExcelBatch batch, string chartName,
        ChartTextElement element, string? fontName = null, double? fontSize = null,
        bool? bold = null, bool? italic = null, string? color = null,
        ChartAxisType? axis = null, int? seriesIndex = null)
    {
        axis = ValidateTextTarget(chartName, element, axis, seriesIndex);
        if (fontName is not null && string.IsNullOrWhiteSpace(fontName))
            throw new ArgumentException("fontName must not be blank.", nameof(fontName));
        if (fontSize is { } size && (!double.IsFinite(size) || size is < 1 or > 409))
            throw new ArgumentOutOfRangeException(nameof(fontSize), "fontSize must be finite and between 1 and 409 points.");
        int? rgb = null;
        if (color is not null)
        {
            if (color.Length != 7 || color[0] != '#' ||
                !int.TryParse(color.AsSpan(1), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out _))
                throw new ArgumentException("color must use #RRGGBB.", nameof(color));
            rgb = FormattingHelpers.ParseColor(color);
        }
        if (fontName is null && fontSize is null && bold is null && italic is null && color is null)
            throw new ArgumentException("Supply at least one text formatting setting.");
        return WithTextFormats(batch, chartName, element, axis, seriesIndex, (font, step) =>
        {
            if (fontName is not null) { step("fontName"); font.Name = fontName; }
            if (fontSize.HasValue) { step("fontSize"); font.Size = fontSize.Value; }
            if (bold.HasValue) { step("bold"); font.Bold = bold.Value; }
            if (italic.HasValue) { step("italic"); font.Italic = italic.Value; }
            if (rgb.HasValue) { step("color"); font.Color = rgb.Value; }
        });
    }

    private static ChartAxisType? ValidateTextTarget(string name, ChartTextElement element,
        ChartAxisType? axis, int? seriesIndex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (!Enum.IsDefined(element)) throw new ArgumentOutOfRangeException(nameof(element));
        if (axis.HasValue && !Enum.IsDefined(axis.Value)) throw new ArgumentOutOfRangeException(nameof(axis));
        if (seriesIndex.HasValue && element != ChartTextElement.DataLabels)
            throw new ArgumentException("seriesIndex applies only to DataLabels.", nameof(seriesIndex));
        if (seriesIndex < 0) throw new ArgumentOutOfRangeException(nameof(seriesIndex), "Series indices must be zero (all series) or positive.");
        if (element == ChartTextElement.AxisTitle && !axis.HasValue)
            throw new ArgumentException("axis is required for AxisTitle.", nameof(axis));
        if (element is not (ChartTextElement.AxisTitle or ChartTextElement.CategoryAxis or ChartTextElement.ValueAxis))
        {
            if (axis.HasValue) throw new ArgumentException("axis applies only to axis text.", nameof(axis));
            return null;
        }
        axis ??= element == ChartTextElement.CategoryAxis ? ChartAxisType.Category : ChartAxisType.Value;
        var (type, group) = MapAxisType(axis.Value);
        if (element == ChartTextElement.CategoryAxis && type != Excel.XlAxisType.xlCategory ||
            element == ChartTextElement.ValueAxis && type != Excel.XlAxisType.xlValue)
            throw new ArgumentException("axis must match the selected category or value element.", nameof(axis));
        return (type, group) switch
        {
            (Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary) => ChartAxisType.Category,
            (Excel.XlAxisType.xlCategory, _) => ChartAxisType.CategorySecondary,
            (Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary) => ChartAxisType.Value,
            _ => ChartAxisType.ValueSecondary
        };
    }

    private static ChartTextFormatResult WithTextFormats(IExcelBatch batch, string name, ChartTextElement element,
        ChartAxisType? axis, int? seriesIndex, Action<Excel.Font, Action<string>>? write)
    {
        return batch.Execute((ctx, ct) =>
        {
            var found = FindChart(ctx.Book, name);
            Excel.SeriesCollection? series = null;
            var completed = new List<string>();
            var target = element.ToString();
            var step = "locating existing element";
            try
            {
                var chart = found.Chart ?? throw new ArgumentException($"Chart '{name}' not found.");
                var result = new ChartTextFormatResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    ChartName = name,
                    Element = element
                };
                int start = 0, end = 0;
                if (element == ChartTextElement.DataLabels)
                {
                    series = (Excel.SeriesCollection)chart.SeriesCollection();
                    if (series.Count == 0) throw new InvalidOperationException($"Chart '{name}' has no data series.");
                    if (seriesIndex > series.Count) throw new ArgumentOutOfRangeException(nameof(seriesIndex), $"Chart has {series.Count} series.");
                    start = seriesIndex is > 0 ? seriesIndex.Value : 1;
                    end = seriesIndex is > 0 ? seriesIndex.Value : series.Count;
                    for (int i = start; i <= end; i++)
                    {
                        ct.ThrowIfCancellationRequested();
                        target = $"{element} series {i}";
                        WithTextFont(chart, element, axis, series, i, ct, (_, _) => { });
                    }
                }
                for (int i = start; i <= end; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    target = i > 0 ? $"{element} series {i}" : $"{element}{(axis.HasValue ? $" {axis}" : "")}";
                    WithTextFont(chart, element, axis, series, i, ct, (font, read) =>
                    {
                        write?.Invoke(font, field => step = $"setting {field}");
                        step = "reading native formatting";
                        var format = read();
                        format.Axis = axis;
                        format.SeriesIndex = i > 0 ? i : null;
                        result.Formats.Add(format);
                    });
                    if (write is not null) completed.Add(target);
                }
                return result;
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    $"Chart '{name}', {target}: Excel failed while {step}. " +
                    (write is null ? "" : $"Completed targets: {string.Join(", ", completed)}. Earlier settings on the current target may remain; no rollback was attempted. ") +
                    ex.Message, ex);
            }
            finally
            {
                ComUtilities.Release(ref series);
                if (found.Chart != null) ComUtilities.Release(ref found.Chart!);
                if (found.Shape != null) ComUtilities.Release(ref found.Shape!);
            }
        });
    }

    private static void WithTextFont(Excel.Chart chart, ChartTextElement element, ChartAxisType? axis,
        Excel.SeriesCollection? seriesCollection, int seriesIndex, CancellationToken ct,
        Action<Excel.Font, Func<ChartTextFormat>> action)
    {
        Excel.ChartArea? area = null;
        Excel.ChartTitle? title = null;
        Excel.Axis? selectedAxis = null;
        Excel.TickLabels? ticks = null;
        Excel.AxisTitle? axisTitle = null;
        Excel.Legend? legend = null;
        Excel.Series? series = null;
        Excel.DataLabels? labels = null;
        Excel.Font? font = null;
        try
        {
            switch (element)
            {
                case ChartTextElement.ChartArea:
                    area = chart.ChartArea;
                    font = area.Font;
                    break;
                case ChartTextElement.Title:
                    if (!chart.HasTitle) throw new InvalidOperationException("The chart has no title. Use set-title first.");
                    title = chart.ChartTitle;
                    font = title.Font;
                    break;
                case ChartTextElement.Legend:
                    if (!chart.HasLegend) throw new InvalidOperationException("The chart has no legend. Use show-legend first.");
                    legend = chart.Legend;
                    font = legend.Font;
                    break;
                case ChartTextElement.DataLabels:
                    series = seriesCollection!.Item(seriesIndex);
                    if (!series.HasDataLabels) throw new InvalidOperationException($"Series {seriesIndex} has no data labels. Use set-data-labels first.");
                    labels = (Excel.DataLabels)series.DataLabels();
                    if (labels.Count == 0) throw new InvalidOperationException($"Series {seriesIndex} has no data labels.");
                    font = labels.Font;
                    break;
                default:
                    var (type, group) = MapAxisType(axis!.Value);
                    if (!(bool)(object)chart.HasAxis[type, group])
                        throw new InvalidOperationException($"The chart has no {axis} axis.");
                    selectedAxis = (Excel.Axis)chart.Axes(type, group);
                    if (element == ChartTextElement.AxisTitle)
                    {
                        if (!selectedAxis.HasTitle) throw new InvalidOperationException($"The {axis} axis has no title. Use set-axis-title first.");
                        axisTitle = selectedAxis.AxisTitle;
                        font = axisTitle.Font;
                    }
                    else
                    {
                        ticks = selectedAxis.TickLabels;
                        font = ticks.Font;
                    }
                    break;
            }
            action(font, () =>
            {
                if (title is not null)
                    return ReadCharacterFormats(font, title.Text.Length, index => title.Characters[index, 1], ct);
                if (axisTitle is not null)
                    return ReadCharacterFormats(font, axisTitle.Text.Length, index => axisTitle.Characters[index, 1], ct);
                if (labels is not null)
                    return ReadDataLabelFonts(labels, ct);
                if (legend is not null)
                    return ReadLegendFonts(legend, font, ct);
                return ReadTextFont(font);
            });
        }
        finally
        {
            ComUtilities.Release(ref font);
            ComUtilities.Release(ref labels);
            ComUtilities.Release(ref series);
            ComUtilities.Release(ref legend);
            ComUtilities.Release(ref axisTitle);
            ComUtilities.Release(ref ticks);
            ComUtilities.Release(ref selectedAxis);
            ComUtilities.Release(ref title);
            ComUtilities.Release(ref area);
        }
    }

    private static ChartTextFormat ReadCharacterFormats(Excel.Font wholeFont, int length,
        Func<int, Excel.Characters> characterAt, CancellationToken ct)
    {
        ChartTextFormat? result = null;
        // Excel's aggregate Font can report the first character rather than mixed formatting.
        for (int i = 1; i <= length; i++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.Characters? character = null;
            Excel.Font? font = null;
            try
            {
                character = characterAt(i);
                font = character.Font;
                result = MergeTextFormats(result, ReadTextFont(font));
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref character);
            }
        }
        return result ?? ReadTextFont(wholeFont);
    }

    private static ChartTextFormat ReadDataLabelFonts(Excel.DataLabels labels, CancellationToken ct)
    {
        ChartTextFormat? result = null;
        for (int i = 1; i <= labels.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            Excel.DataLabel? label = null;
            Excel.Font? font = null;
            try
            {
                label = labels.Item(i);
                font = label.Font;
                // Excel-generated label text has no editable character range.
                result = MergeTextFormats(result, label.AutoText
                    ? ReadTextFont(font)
                    : ReadCharacterFormats(font, label.Text.Length, index => label.Characters[index, 1], ct));
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref label);
            }
        }
        return result ?? throw new InvalidOperationException("The selected series has no data labels.");
    }

    private static ChartTextFormat ReadLegendFonts(Excel.Legend legend, Excel.Font wholeFont, CancellationToken ct)
    {
        Excel.LegendEntries? entries = null;
        ChartTextFormat? result = null;
        try
        {
            entries = (Excel.LegendEntries)legend.LegendEntries();
            for (int i = 1; i <= entries.Count; i++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.LegendEntry? entry = null;
                Excel.Font? font = null;
                try
                {
                    entry = entries.Item(i);
                    font = entry.Font;
                    result = MergeTextFormats(result, ReadTextFont(font));
                }
                finally
                {
                    ComUtilities.Release(ref font);
                    ComUtilities.Release(ref entry);
                }
            }
            return result ?? ReadTextFont(wholeFont);
        }
        finally { ComUtilities.Release(ref entries); }
    }

    private static ChartTextFormat MergeTextFormats(ChartTextFormat? left, ChartTextFormat right)
    {
        if (left is null) return right;
        if (left.FontName != right.FontName) left.FontName = null;
        if (left.FontSize != right.FontSize) left.FontSize = null;
        if (left.Bold != right.Bold) left.Bold = null;
        if (left.Italic != right.Italic) left.Italic = null;
        if (left.Color != right.Color) left.Color = null;
        SetMixedTextFields(left);
        return left;
    }

    private static ChartTextFormat ReadTextFont(Excel.Font font)
    {
        var result = new ChartTextFormat();
        object? name = font.Name, size = font.Size, bold = font.Bold, italic = font.Italic, color = font.Color;
        result.FontName = name is null or DBNull ? null : Convert.ToString(name, CultureInfo.InvariantCulture);
        result.FontSize = size is null or DBNull ? null : Convert.ToDouble(size, CultureInfo.InvariantCulture);
        result.Bold = bold is null or DBNull ? null : Convert.ToBoolean(bold, CultureInfo.InvariantCulture);
        result.Italic = italic is null or DBNull ? null : Convert.ToBoolean(italic, CultureInfo.InvariantCulture);
        result.Color = color is null or DBNull ? null : FormattingHelpers.ColorToHex(Convert.ToInt32(color, CultureInfo.InvariantCulture));
        SetMixedTextFields(result);
        return result;
    }

    private static void SetMixedTextFields(ChartTextFormat result)
    {
        result.MixedFields.Clear();
        if (result.FontName is null) result.MixedFields.Add("fontName");
        if (result.FontSize is null) result.MixedFields.Add("fontSize");
        if (result.Bold is null) result.MixedFields.Add("bold");
        if (result.Italic is null) result.MixedFields.Add("italic");
        if (result.Color is null) result.MixedFields.Add("color");
    }
}
