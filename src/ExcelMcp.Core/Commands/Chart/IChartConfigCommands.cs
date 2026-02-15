using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chart configuration - data source, series, type, title, axis labels, legend, and styling.
///
/// SERIES MANAGEMENT:
/// - add-series: Add data series with valuesRange (required) and optional categoryRange
/// - remove-series: Remove series by 1-based index
/// - set-source-range: Replace entire chart data source
///
/// TITLES AND LABELS:
/// - set-title: Set chart title (empty string hides title)
/// - set-axis-title: Set axis labels (Category, Value, CategorySecondary, ValueSecondary)
///
/// CHART STYLES: 1-48 (built-in Excel styles with different color schemes)
///
/// DATA LABELS: Show values, percentages, series/category names.
/// Positions: Center, InsideEnd, InsideBase, OutsideEnd, BestFit.
///
/// TRENDLINES: Linear, Exponential, Logarithmic, Polynomial (order 2-6), Power, MovingAverage.
///
/// PLACEMENT MODE:
/// - 1: Move and size with cells
/// - 2: Move but don't size with cells
/// - 3: Don't move or size with cells (free floating)
///
/// Use chart for lifecycle operations (create, delete, move, fit-to-range).
/// </summary>
[ServiceCategory("chartconfig", "ChartConfig")]
[McpTool("chart_config", Title = "Chart Configuration", Destructive = true, Category = "analysis",
    Description = "Chart configuration - data source, series, type, title, axis labels, legend, and styling. SERIES: add-series (valuesRange required), remove-series (1-based index), set-source-range. TITLES: set-title, set-axis-title (Category/Value/Secondary). AXIS: number format, scale min/max/units. LEGEND: Bottom, Corner, Top, Right, Left. STYLES: 1-48 built-in. DATA LABELS: values, percentages, positions (Center, InsideEnd, OutsideEnd, BestFit). GRIDLINES: major/minor for value/category axes. TRENDLINES: Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage. SERIES FORMAT: marker style/size/colors, invert if negative. PLACEMENT: 1=move+size with cells, 2=move only, 3=free floating. Use chart for lifecycle.")]
public interface IChartConfigCommands
{
    // === DATA SOURCE OPERATIONS ===

    /// <summary>
    /// Sets data source range for Regular Charts.
    /// PivotCharts: Throws exception guiding to pivottable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="sourceRange">New data source range (e.g., Sheet1!A1:D10)</param>
    [ServiceAction("set-source-range")]
    OperationResult SetSourceRange(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] string sourceRange);

    /// <summary>
    /// Adds a data series to Regular Charts.
    /// PivotCharts: Throws exception guiding to pivottable(action: 'add-value-field').
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesName">Display name for the series</param>
    /// <param name="valuesRange">Range containing series values (e.g., B2:B10)</param>
    /// <param name="categoryRange">Optional range for category labels (e.g., A2:A10)</param>
    [ServiceAction("add-series")]
    SeriesInfo AddSeries(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] string seriesName,
        [RequiredParameter] string valuesRange,
        string? categoryRange = null);

    /// <summary>
    /// Removes a data series from Regular Charts.
    /// PivotCharts: Throws exception guiding to pivottable(action: 'remove-field').
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of the series to remove</param>
    [ServiceAction("remove-series")]
    OperationResult RemoveSeries(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex);

    // === APPEARANCE OPERATIONS ===

    /// <summary>
    /// Changes chart type (works for both Regular and PivotCharts).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="chartType">New chart type to apply</param>
    [ServiceAction("set-chart-type")]
    OperationResult SetChartType(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartType chartType);

    /// <summary>
    /// Sets chart title.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="title">Title text to display</param>
    [ServiceAction("set-title")]
    OperationResult SetTitle(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] string title);

    /// <summary>
    /// Sets axis title.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Which axis to set title for (Category, Value, SeriesAxis)</param>
    /// <param name="title">Axis title text</param>
    [ServiceAction("set-axis-title")]
    OperationResult SetAxisTitle(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis,
        [RequiredParameter] string title);

    /// <summary>
    /// Gets axis number format for tick labels.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Which axis to get format from</param>
    [ServiceAction("get-axis-number-format")]
    string GetAxisNumberFormat(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis);

    /// <summary>
    /// Sets axis number format for tick labels.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Which axis to format</param>
    /// <param name="numberFormat">Excel number format code (e.g., "$#,##0", "0.00%")</param>
    [ServiceAction("set-axis-number-format")]
    OperationResult SetAxisNumberFormat(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis,
        [RequiredParameter] string numberFormat);

    /// <summary>
    /// Shows or hides chart legend.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="visible">True to show legend, false to hide</param>
    /// <param name="legendPosition">Optional position for the legend</param>
    [ServiceAction("show-legend")]
    OperationResult ShowLegend(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] bool visible,
        LegendPosition? legendPosition = null);

    /// <summary>
    /// Applies a chart style.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="styleId">Excel chart style ID (1-48 for most chart types)</param>
    [ServiceAction("set-style")]
    OperationResult SetStyle(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int styleId);

    /// <summary>
    /// Sets chart placement mode (how chart responds when underlying cells are resized).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="placement">Placement mode: 1=MoveAndSize, 2=Move, 3=FreeFloating</param>
    [ServiceAction("set-placement")]
    OperationResult SetPlacement(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int placement);

    // === DATA LABELS ===

    /// <summary>
    /// Configures data labels for chart series.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="showValue">Show data values on labels</param>
    /// <param name="showPercentage">Show percentage values (pie/doughnut charts)</param>
    /// <param name="showSeriesName">Show series name on labels</param>
    /// <param name="showCategoryName">Show category name on labels</param>
    /// <param name="showBubbleSize">Show bubble size (bubble charts)</param>
    /// <param name="separator">Separator string between label components</param>
    /// <param name="labelPosition">Position of data labels relative to data points</param>
    /// <param name="seriesIndex">Optional 1-based series index (null for all series)</param>
    [ServiceAction("set-data-labels")]
    OperationResult SetDataLabels(
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Which axis to get scale settings from</param>
    [ServiceAction("get-axis-scale")]
    AxisScaleResult GetAxisScale(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis);

    /// <summary>
    /// Sets axis scale settings.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Which axis to configure</param>
    /// <param name="minimumScale">Minimum axis value (null for auto)</param>
    /// <param name="maximumScale">Maximum axis value (null for auto)</param>
    /// <param name="majorUnit">Major gridline interval (null for auto)</param>
    /// <param name="minorUnit">Minor gridline interval (null for auto)</param>
    [ServiceAction("set-axis-scale")]
    OperationResult SetAxisScale(
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    [ServiceAction("get-gridlines")]
    GridlinesResult GetGridlines(
        IExcelBatch batch,
        [RequiredParameter] string chartName);

    /// <summary>
    /// Configures gridlines visibility for chart axes.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Which axis gridlines to configure</param>
    /// <param name="showMajor">Show major gridlines (null to keep current)</param>
    /// <param name="showMinor">Show minor gridlines (null to keep current)</param>
    [ServiceAction("set-gridlines")]
    OperationResult SetGridlines(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] ChartAxisType axis,
        bool? showMajor = null,
        bool? showMinor = null);

    // === SERIES FORMATTING ===

    /// <summary>
    /// Configures series marker formatting (for line, scatter, and radar charts).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of the series</param>
    /// <param name="markerStyle">Marker shape style</param>
    /// <param name="markerSize">Marker size in points (2-72)</param>
    /// <param name="markerBackgroundColor">Marker fill color (#RRGGBB)</param>
    /// <param name="markerForegroundColor">Marker border color (#RRGGBB)</param>
    /// <param name="invertIfNegative">Invert colors for negative values</param>
    [ServiceAction("set-series-format")]
    OperationResult SetSeriesFormat(
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of the series</param>
    [ServiceAction("list-trendlines")]
    TrendlineListResult ListTrendlines(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex);

    /// <summary>
    /// Adds a trendline to a chart series.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of the series</param>
    /// <param name="type">Type of trendline (Linear, Exponential, etc.)</param>
    /// <param name="order">Polynomial order (2-6, for Polynomial type)</param>
    /// <param name="period">Moving average period (for MovingAverage type)</param>
    /// <param name="forward">Periods to extend forward</param>
    /// <param name="backward">Periods to extend backward</param>
    /// <param name="intercept">Force trendline through specific Y-intercept</param>
    /// <param name="displayEquation">Display trendline equation on chart</param>
    /// <param name="displayRSquared">Display R-squared value on chart</param>
    /// <param name="name">Custom name for the trendline</param>
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
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of the series</param>
    /// <param name="trendlineIndex">1-based index of the trendline to delete</param>
    [ServiceAction("delete-trendline")]
    OperationResult DeleteTrendline(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] int seriesIndex,
        [RequiredParameter] int trendlineIndex);

    /// <summary>
    /// Updates trendline properties.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of the series</param>
    /// <param name="trendlineIndex">1-based index of the trendline</param>
    /// <param name="forward">Periods to extend forward (null to keep current)</param>
    /// <param name="backward">Periods to extend backward (null to keep current)</param>
    /// <param name="intercept">Force through Y-intercept (null to keep current)</param>
    /// <param name="displayEquation">Display equation (null to keep current)</param>
    /// <param name="displayRSquared">Display R-squared (null to keep current)</param>
    /// <param name="name">Custom name (null to keep current)</param>
    [ServiceAction("set-trendline")]
    OperationResult SetTrendline(
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
