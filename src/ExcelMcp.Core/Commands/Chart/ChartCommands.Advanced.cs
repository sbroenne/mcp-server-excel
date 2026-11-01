using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Advanced chart plotting, combo-series, and material formatting operations.
/// </summary>
public partial class ChartCommands
{
    /// <inheritdoc />
    public OperationResult SetSeriesChartType(
        IExcelBatch batch,
        string chartName,
        int seriesIndex,
        ChartType chartType)
    {
        ValidateSeriesIndex(seriesIndex);
        if (!Enum.IsDefined(chartType)) throw new ArgumentOutOfRangeException(nameof(chartType));
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            Excel.Chart? chart = null;
            Excel.SeriesCollection? seriesCollection = null;
            Excel.Series? series = null;
            try
            {
                chart = findResult.Chart;
                findResult.Chart = null;
                if (_pivotStrategy.CanHandle(chart))
                {
                    throw new InvalidOperationException(
                        "Per-series chart types are not supported for PivotCharts. " +
                        "Use set-chart-type to change the entire PivotChart.");
                }

                seriesCollection = (Excel.SeriesCollection)chart.SeriesCollection();
                var seriesCount = seriesCollection.Count;
                if (seriesIndex < 1 || seriesIndex > seriesCount)
                {
                    throw new ArgumentException(
                        $"Series index {seriesIndex} is out of range. Chart has {seriesCount} series.",
                        nameof(seriesIndex));
                }

                series = seriesCollection.Item(seriesIndex);
                series.ChartType = (Excel.XlChartType)chartType;

                return new OperationResult
                {
                    Success = true,
                    Action = "set-series-chart-type",
                    Message = $"Series {seriesIndex} now uses {chartType}.",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref series);
                ComUtilities.Release(ref seriesCollection);
                ComUtilities.Release(ref chart);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public ChartPlotOptionsResult GetPlotOptions(IExcelBatch batch, string chartName)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            Excel.Chart? chart = null;
            try
            {
                chart = findResult.Chart;
                findResult.Chart = null;
                return new ChartPlotOptionsResult
                {
                    Success = true,
                    ChartName = chartName,
                    PlotBy = (ChartPlotBy)chart.PlotBy,
                    DisplayBlanksAs = (ChartDisplayBlanksAs)chart.DisplayBlanksAs,
                    PlotVisibleOnly = chart.PlotVisibleOnly,
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref chart);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetPlotOptions(
        IExcelBatch batch,
        string chartName,
        ChartPlotBy? plotBy = null,
        ChartDisplayBlanksAs? displayBlanksAs = null,
        bool? plotVisibleOnly = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            Excel.Chart? chart = null;
            try
            {
                chart = findResult.Chart;
                findResult.Chart = null;
                if (plotBy.HasValue)
                {
                    chart.PlotBy = (Excel.XlRowCol)plotBy.Value;
                }

                if (displayBlanksAs.HasValue)
                {
                    chart.DisplayBlanksAs = (Excel.XlDisplayBlanksAs)displayBlanksAs.Value;
                }

                if (plotVisibleOnly.HasValue)
                {
                    chart.PlotVisibleOnly = plotVisibleOnly.Value;
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "set-plot-options",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref chart);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetAreaFormat(
        IExcelBatch batch,
        string chartName,
        ChartAreaTarget area,
        string? fillColor = null,
        double? fillTransparency = null,
        string? lineColor = null,
        double? lineWeight = null,
        bool? lineVisible = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(chartName);
        if (!Enum.IsDefined(area)) throw new ArgumentOutOfRangeException(nameof(area));
        ValidateMaterialFormat(fillTransparency, lineWeight);
        if (fillColor is not null) _ = FormattingHelpers.ParseColor(fillColor);
        if (lineColor is not null) _ = FormattingHelpers.ParseColor(lineColor);
        return batch.Execute((ctx, ct) =>
        {
            var findResult = FindChart(ctx.Book, chartName);
            if (findResult.Chart == null)
            {
                throw new InvalidOperationException($"Chart '{chartName}' not found in workbook.");
            }

            Excel.Chart? chart = null;
            Excel.ChartArea? chartArea = null;
            Excel.PlotArea? plotArea = null;
            Excel.ChartFormat? format = null;
            try
            {
                chart = findResult.Chart;
                findResult.Chart = null;
                if (area == ChartAreaTarget.Chart)
                {
                    chartArea = chart.ChartArea;
                    format = chartArea.Format;
                }
                else
                {
                    plotArea = chart.PlotArea;
                    format = plotArea.Format;
                }

                ApplyMaterialFormat(format, fillColor, fillTransparency, lineColor, lineWeight, lineVisible);
                return new OperationResult
                {
                    Success = true,
                    Action = "set-area-format",
                    FilePath = batch.WorkbookPath
                };
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException(
                    $"Chart '{chartName}', {area} area: Excel failed while formatting the fill or border. " +
                    $"The border may be visible and earlier fill, colour or weight changes may remain; no rollback was attempted. {ex.Message}", ex);
            }
            finally
            {
                ComUtilities.Release(ref format!);
                ComUtilities.Release(ref plotArea);
                ComUtilities.Release(ref chartArea);
                ComUtilities.Release(ref chart);
                if (findResult.Shape != null) ComUtilities.Release(ref findResult.Shape!);
                if (findResult.Chart != null) ComUtilities.Release(ref findResult.Chart!);
            }
        });
    }

    private static void ApplyMaterialFormat(
        Excel.ChartFormat format,
        string? fillColor,
        double? fillTransparency,
        string? lineColor,
        double? lineWeight,
        bool? lineVisible = null)
    {
        ValidateMaterialFormat(fillTransparency, lineWeight);

        Excel.FillFormat? fill = null;
        Excel.ColorFormat? fillForeColor = null;
        try
        {
            if (fillColor != null || fillTransparency.HasValue)
            {
                fill = format.Fill;
                fill.Solid();
                if (fillColor != null)
                {
                    fillForeColor = fill.ForeColor;
                    fillForeColor.RGB = FormattingHelpers.ParseColor(fillColor);
                }

                if (fillTransparency.HasValue)
                {
                    fill.Transparency = (float)fillTransparency.Value;
                }
            }

            ApplyLineFormat(format, lineColor, lineWeight,
                lineVisible ?? (lineColor is not null || lineWeight.HasValue ? true : null));
        }
        finally
        {
            ComUtilities.Release(ref fillForeColor!);
            ComUtilities.Release(ref fill!);
        }
    }

    private static void ApplyLineFormat(Excel.ChartFormat format, string? lineColor,
        double? lineWeight, bool? lineVisible)
    {
        if (lineColor is null && lineWeight is null && lineVisible is null) return;
        Excel.LineFormat? line = null;
        Excel.ColorFormat? color = null;
        try
        {
            line = format.Line;
            var visibility = lineVisible.HasValue
                ? (lineVisible.Value ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse)
                : line.Visible;
            // Excel ignores colour writes on hidden lines until the line is shown.
            if (lineColor is not null || lineWeight.HasValue) line.Visible = Office.MsoTriState.msoTrue;
            if (lineColor is not null)
            {
                color = line.ForeColor;
                color.RGB = FormattingHelpers.ParseColor(lineColor);
            }
            if (lineWeight.HasValue) line.Weight = (float)lineWeight.Value;
            line.Visible = visibility;
        }
        finally
        {
            ComUtilities.Release(ref color);
            ComUtilities.Release(ref line);
        }
    }

    private static void ValidateMaterialFormat(double? fillTransparency, double? lineWeight)
    {
        if (fillTransparency.HasValue && (!double.IsFinite(fillTransparency.Value) || fillTransparency is < 0 or > 1))
        {
            throw new ArgumentOutOfRangeException(
                nameof(fillTransparency),
                fillTransparency,
                "Fill transparency must be finite and between 0 and 1.");
        }

        if (lineWeight.HasValue && (!double.IsFinite(lineWeight.Value) || !float.IsFinite((float)lineWeight.Value) || lineWeight is <= 0))
        {
            throw new ArgumentOutOfRangeException(
                nameof(lineWeight),
                lineWeight,
                "Line weight must be finite and greater than zero.");
        }
    }
}
