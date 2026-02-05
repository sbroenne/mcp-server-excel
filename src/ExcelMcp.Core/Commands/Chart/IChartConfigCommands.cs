using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Excel chart configuration operations - data sources, titles, axes, styling, trendlines.
/// Lifecycle operations (create, delete, move) are in IChartCommands.
/// </summary>
[ServiceCategory("chartconfig", "ChartConfig")]
[McpTool("excel_chart_config")]
public interface IChartConfigCommands
{
    // === DATA SOURCE OPERATIONS ===

    /// <summary>
    /// Sets data source range for Regular Charts.
    /// PivotCharts: Throws exception guiding to excel_pivottable.
    /// </summary>
    [ServiceAction("set-source-range")]
    void SetSourceRange(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] string sourceRange);

    /// <summary>
    /// Adds a data series to Regular Charts.
    /// PivotCharts: Throws exception guiding to excel_pivottable(action: 'add-value-field').
    /// </summary>
    [ServiceAction("add-series")]
    SeriesInfo AddSeries(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] string seriesName,
        [RequiredParameter] string valuesRange,
        string? categoryRange = null);

    /// <summary>
    /// Removes a data series from Regular Charts.
    /// PivotCharts: Throws exception guiding to excel_pivottable(action: 'remove-field').
    /// </summary>
    [ServiceAction("remove-series")]
    void RemoveSeries(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex);

    // === APPEARANCE OPERATIONS ===

    /// <summary>
    /// Changes chart type (works for both Regular and PivotCharts).
    /// </summary>
    [ServiceAction("set-chart-type")]
    void SetChartType(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartType chartType);

    /// <summary>
    /// Sets chart title.
    /// </summary>
    [ServiceAction("set-title")]
    void SetTitle(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] string title);

    /// <summary>
    /// Sets axis title.
    /// </summary>
    [ServiceAction("set-axis-title")]
    void SetAxisTitle(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis,
        [RequiredParameter] string title);

    /// <summary>
    /// Gets axis number format for tick labels.
    /// </summary>
    [ServiceAction("get-axis-number-format")]
    string GetAxisNumberFormat(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis);

    /// <summary>
    /// Sets axis number format for tick labels.
    /// </summary>
    [ServiceAction("set-axis-number-format")]
    void SetAxisNumberFormat(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis,
        [RequiredParameter] string numberFormat);

    /// <summary>
    /// Shows or hides chart legend.
    /// </summary>
    [ServiceAction("show-legend")]
    void ShowLegend(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] bool visible,
        LegendPosition? legendPosition = null);

    /// <summary>
    /// Applies a chart style.
    /// </summary>
    [ServiceAction("set-style")]
    void SetStyle(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int styleId);

    /// <summary>
    /// Sets chart placement mode (how chart responds when underlying cells are resized).
    /// </summary>
    [ServiceAction("set-placement")]
    void SetPlacement(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int placement);

    // === DATA LABELS ===

    /// <summary>
    /// Configures data labels for chart series.
    /// </summary>
    [ServiceAction("set-data-labels")]
    void SetDataLabels(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        bool? showValue = null,
        bool? showPercentage = null,
        bool? showSeriesName = null,
        bool? showCategoryName = null,
        bool? showBubbleSize = null,
        string? separator = null,
        DataLabelPosition? labelPosition = null,
        int? seriesIndex = null);

    // === AXIS SCALE ===

    /// <summary>
    /// Gets axis scale settings.
    /// </summary>
    [ServiceAction("get-axis-scale")]
    AxisScaleResult GetAxisScale(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis);

    /// <summary>
    /// Sets axis scale settings.
    /// </summary>
    [ServiceAction("set-axis-scale")]
    void SetAxisScale(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis,
        double? minimumScale = null,
        double? maximumScale = null,
        double? majorUnit = null,
        double? minorUnit = null);

    // === GRIDLINES ===

    /// <summary>
    /// Gets gridlines visibility for chart axes.
    /// </summary>
    [ServiceAction("get-gridlines")]
    GridlinesResult GetGridlines(
        IExcelBatch batch,
        [RequiredParameter] string chartName);

    /// <summary>
    /// Configures gridlines visibility for chart axes.
    /// </summary>
    [ServiceAction("set-gridlines")]
    void SetGridlines(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis,
        bool? showMajor = null,
        bool? showMinor = null);

    // === SERIES FORMATTING ===

    /// <summary>
    /// Configures series marker formatting (for line, scatter, and radar charts).
    /// </summary>
    [ServiceAction("set-series-format")]
    void SetSeriesFormat(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex,
        MarkerStyle? markerStyle = null,
        int? markerSize = null,
        string? markerBackgroundColor = null,
        string? markerForegroundColor = null,
        bool? invertIfNegative = null);

    // === TRENDLINES ===

    /// <summary>
    /// Lists all trendlines on a chart series.
    /// </summary>
    [ServiceAction("list-trendlines")]
    TrendlineListResult ListTrendlines(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex);

    /// <summary>
    /// Adds a trendline to a chart series.
    /// </summary>
    [ServiceAction("add-trendline")]
    TrendlineResult AddTrendline(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex,
        [RequiredParameter] TrendlineType type,
        int? order = null,
        int? period = null,
        double? forward = null,
        double? backward = null,
        double? intercept = null,
        bool displayEquation = false,
        bool displayRSquared = false,
        string? name = null);

    /// <summary>
    /// Deletes a trendline from a chart series.
    /// </summary>
    [ServiceAction("delete-trendline")]
    void DeleteTrendline(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex,
        [RequiredParameter] int trendlineIndex);

    /// <summary>
    /// Updates trendline properties.
    /// </summary>
    [ServiceAction("set-trendline")]
    void SetTrendline(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex,
        [RequiredParameter] int trendlineIndex,
        double? forward = null,
        double? backward = null,
        double? intercept = null,
        bool? displayEquation = null,
        bool? displayRSquared = null,
        string? name = null);
}
