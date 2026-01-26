using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Excel chart operations - creating and managing Regular Charts and PivotCharts.
/// Supports two chart types: Regular (static, from ranges) and PivotCharts (dynamic, from PivotTables).
/// </summary>
public interface IChartCommands
{
    // === LIFECYCLE OPERATIONS ===

    /// <summary>
    /// Lists all charts in workbook (Regular and PivotCharts).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>List of charts with names, types, sheets, positions, data sources</returns>
    List<ChartInfo> List(IExcelBatch batch);

    /// <summary>
    /// Gets complete chart configuration.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart (or shape name)</param>
    /// <returns>Chart type, data source, series info, position, styling</returns>
    ChartInfoResult Read(IExcelBatch batch, string chartName);

    /// <summary>
    /// Creates a Regular Chart from an Excel range.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name for chart placement</param>
    /// <param name="sourceRange">Source data range (e.g., "A1:D10")</param>
    /// <param name="chartType">Chart type from ChartType enum</param>
    /// <param name="left">Left position in points</param>
    /// <param name="top">Top position in points</param>
    /// <param name="width">Width in points (default: 400)</param>
    /// <param name="height">Height in points (default: 300)</param>
    /// <param name="chartName">Optional name for the chart</param>
    /// <returns>Created chart name and configuration</returns>
    ChartCreateResult CreateFromRange(
        IExcelBatch batch,
        string sheetName,
        string sourceRange,
        ChartType chartType,
        double left,
        double top,
        double width = 400,
        double height = 300,
        string? chartName = null);

    /// <summary>
    /// Creates a PivotChart from an existing PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="sheetName">Worksheet name for chart placement</param>
    /// <param name="chartType">Chart type from ChartType enum</param>
    /// <param name="left">Left position in points</param>
    /// <param name="top">Top position in points</param>
    /// <param name="width">Width in points (default: 400)</param>
    /// <param name="height">Height in points (default: 300)</param>
    /// <param name="chartName">Optional name for the chart</param>
    /// <returns>Created PivotChart name and linked PivotTable</returns>
    ChartCreateResult CreateFromPivotTable(
        IExcelBatch batch,
        string pivotTableName,
        string sheetName,
        ChartType chartType,
        double left,
        double top,
        double width = 400,
        double height = 300,
        string? chartName = null);

    /// <summary>
    /// Deletes a chart (Regular or PivotChart).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart to delete</param>
    void Delete(IExcelBatch batch, string chartName);

    /// <summary>
    /// Moves/resizes a chart.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="left">New left position in points (null = no change)</param>
    /// <param name="top">New top position in points (null = no change)</param>
    /// <param name="width">New width in points (null = no change)</param>
    /// <param name="height">New height in points (null = no change)</param>
    void Move(
        IExcelBatch batch,
        string chartName,
        double? left = null,
        double? top = null,
        double? width = null,
        double? height = null);

    // === DATA SOURCE OPERATIONS ===

    /// <summary>
    /// Sets data source range for Regular Charts.
    /// PivotCharts: Throws exception guiding to excel_pivottable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="sourceRange">New source range (e.g., "Sheet1!A1:D10")</param>
    void SetSourceRange(IExcelBatch batch, string chartName, string sourceRange);

    /// <summary>
    /// Adds a data series to Regular Charts.
    /// PivotCharts: Throws exception guiding to excel_pivottable(action: 'add-value-field').
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesName">Name for the series</param>
    /// <param name="valuesRange">Range containing Y values (e.g., "Sheet1!B2:B10")</param>
    /// <param name="categoryRange">Optional range for X values/categories</param>
    /// <returns>Series information</returns>
    SeriesInfo AddSeries(
        IExcelBatch batch,
        string chartName,
        string seriesName,
        string valuesRange,
        string? categoryRange = null);

    /// <summary>
    /// Removes a data series from Regular Charts.
    /// PivotCharts: Throws exception guiding to excel_pivottable(action: 'remove-field').
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of series to remove</param>
    void RemoveSeries(IExcelBatch batch, string chartName, int seriesIndex);

    // === APPEARANCE OPERATIONS ===

    /// <summary>
    /// Changes chart type (works for both Regular and PivotCharts).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="chartType">New chart type from ChartType enum</param>
    void SetChartType(IExcelBatch batch, string chartName, ChartType chartType);

    /// <summary>
    /// Sets chart title.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="title">Chart title text (empty to hide title)</param>
    void SetTitle(IExcelBatch batch, string chartName, string title);

    /// <summary>
    /// Sets axis title.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Axis type (Primary, Secondary, Category, Value)</param>
    /// <param name="title">Axis title text</param>
    void SetAxisTitle(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        string title);

    /// <summary>
    /// Gets axis number format for tick labels.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Axis type (Primary, Secondary, Category, Value)</param>
    /// <returns>The current number format string</returns>
    string GetAxisNumberFormat(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis);

    /// <summary>
    /// Sets axis number format for tick labels.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Axis type (Primary, Secondary, Category, Value)</param>
    /// <param name="numberFormat">Excel number format string (e.g., "$#,##0,,"M"" for millions)</param>
    void SetAxisNumberFormat(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        string numberFormat);

    /// <summary>
    /// Shows or hides chart legend.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="visible">True to show legend, false to hide</param>
    /// <param name="position">Legend position (optional)</param>
    void ShowLegend(
        IExcelBatch batch,
        string chartName,
        bool visible,
        LegendPosition? position = null);

    /// <summary>
    /// Applies a chart style.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="styleId">Style number (1-48)</param>
    void SetStyle(IExcelBatch batch, string chartName, int styleId);

    /// <summary>
    /// Sets chart placement mode (how chart responds when underlying cells are resized).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="placement">Placement mode: 1=Move and size with cells, 2=Move but don't size, 3=Don't move or size</param>
    void SetPlacement(IExcelBatch batch, string chartName, int placement);

    /// <summary>
    /// Fits a chart to a cell range by setting position and size to match the range bounds.
    /// Uses the range's Left, Top, Width, Height properties.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="sheetName">Worksheet containing the range</param>
    /// <param name="rangeAddress">Target range (e.g., "A1:D10")</param>
    void FitToRange(IExcelBatch batch, string chartName, string sheetName, string rangeAddress);

    // === DATA LABELS ===

    /// <summary>
    /// Configures data labels for chart series.
    /// Applies to all series, or specify seriesIndex for one series.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="showValue">Show actual data values</param>
    /// <param name="showPercentage">Show percentage (pie/doughnut charts)</param>
    /// <param name="showSeriesName">Show series name in label</param>
    /// <param name="showCategoryName">Show category name in label</param>
    /// <param name="showBubbleSize">Show bubble size (bubble charts)</param>
    /// <param name="separator">Separator between label parts (e.g., ", " or "\n")</param>
    /// <param name="position">Label position (Center, InsideEnd, InsideBase, OutsideEnd, BestFit)</param>
    /// <param name="seriesIndex">Optional 1-based series index (null = all series)</param>
    void SetDataLabels(
        IExcelBatch batch,
        string chartName,
        bool? showValue = null,
        bool? showPercentage = null,
        bool? showSeriesName = null,
        bool? showCategoryName = null,
        bool? showBubbleSize = null,
        string? separator = null,
        DataLabelPosition? position = null,
        int? seriesIndex = null);

    // === AXIS SCALE ===

    /// <summary>
    /// Gets axis scale settings.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Axis type (Value, Category, ValueSecondary, CategorySecondary)</param>
    /// <returns>Axis scale information including min/max/units</returns>
    AxisScaleResult GetAxisScale(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis);

    /// <summary>
    /// Sets axis scale settings.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Axis type (Value, Category, ValueSecondary, CategorySecondary)</param>
    /// <param name="minimumScale">Minimum axis value (null = auto)</param>
    /// <param name="maximumScale">Maximum axis value (null = auto)</param>
    /// <param name="majorUnit">Major unit interval (null = auto)</param>
    /// <param name="minorUnit">Minor unit interval (null = auto)</param>
    void SetAxisScale(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
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
    /// <returns>Gridlines visibility for value and category axes</returns>
    GridlinesResult GetGridlines(IExcelBatch batch, string chartName);

    /// <summary>
    /// Configures gridlines visibility for chart axes.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Axis type (Value, Category)</param>
    /// <param name="showMajor">Show major gridlines</param>
    /// <param name="showMinor">Show minor gridlines</param>
    void SetGridlines(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        bool? showMajor = null,
        bool? showMinor = null);

    // === SERIES FORMATTING ===

    /// <summary>
    /// Configures series marker formatting (for line, scatter, and radar charts).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based series index</param>
    /// <param name="markerStyle">Marker style (None, Square, Diamond, Triangle, X, Star, Circle, Plus, etc.)</param>
    /// <param name="markerSize">Marker size in points (2-72)</param>
    /// <param name="markerBackgroundColor">Marker fill color (#RRGGBB hex)</param>
    /// <param name="markerForegroundColor">Marker border color (#RRGGBB hex)</param>
    /// <param name="invertIfNegative">Invert colors for negative values</param>
    void SetSeriesFormat(
        IExcelBatch batch,
        string chartName,
        int seriesIndex,
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
    /// <param name="seriesIndex">1-based series index</param>
    /// <returns>List of trendlines with their properties</returns>
    TrendlineListResult ListTrendlines(
        IExcelBatch batch,
        string chartName,
        int seriesIndex);

    /// <summary>
    /// Adds a trendline to a chart series.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based series index</param>
    /// <param name="type">Trendline type (Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage)</param>
    /// <param name="order">Polynomial order (2-6), required when type is Polynomial</param>
    /// <param name="period">Moving average period, required when type is MovingAverage</param>
    /// <param name="forward">Periods to forecast forward</param>
    /// <param name="backward">Periods to forecast backward</param>
    /// <param name="intercept">Y-intercept value (null = calculated automatically)</param>
    /// <param name="displayEquation">Display equation on chart</param>
    /// <param name="displayRSquared">Display R-squared value on chart</param>
    /// <param name="name">Custom name for the trendline</param>
    /// <returns>Created trendline information</returns>
    TrendlineResult AddTrendline(
        IExcelBatch batch,
        string chartName,
        int seriesIndex,
        TrendlineType type,
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
    /// <param name="seriesIndex">1-based series index</param>
    /// <param name="trendlineIndex">1-based trendline index within the series</param>
    void DeleteTrendline(
        IExcelBatch batch,
        string chartName,
        int seriesIndex,
        int trendlineIndex);

    /// <summary>
    /// Updates trendline properties.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based series index</param>
    /// <param name="trendlineIndex">1-based trendline index within the series</param>
    /// <param name="forward">Periods to forecast forward</param>
    /// <param name="backward">Periods to forecast backward</param>
    /// <param name="intercept">Y-intercept value (null = calculated automatically)</param>
    /// <param name="displayEquation">Display equation on chart</param>
    /// <param name="displayRSquared">Display R-squared value on chart</param>
    /// <param name="name">Custom name for the trendline</param>
    void SetTrendline(
        IExcelBatch batch,
        string chartName,
        int seriesIndex,
        int trendlineIndex,
        double? forward = null,
        double? backward = null,
        double? intercept = null,
        bool? displayEquation = null,
        bool? displayRSquared = null,
        string? name = null);
}
