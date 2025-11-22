using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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
    ChartListResult List(IExcelBatch batch);

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
    /// <returns>Operation result</returns>
    OperationResult Delete(IExcelBatch batch, string chartName);

    /// <summary>
    /// Moves/resizes a chart.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="left">New left position in points (null = no change)</param>
    /// <param name="top">New top position in points (null = no change)</param>
    /// <param name="width">New width in points (null = no change)</param>
    /// <param name="height">New height in points (null = no change)</param>
    /// <returns>Operation result with new position</returns>
    OperationResult Move(
        IExcelBatch batch,
        string chartName,
        double? left = null,
        double? top = null,
        double? width = null,
        double? height = null);

    // === DATA SOURCE OPERATIONS ===

    /// <summary>
    /// Sets data source range for Regular Charts.
    /// PivotCharts: Returns error guiding to excel_pivottable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="sourceRange">New source range (e.g., "Sheet1!A1:D10")</param>
    /// <returns>Operation result</returns>
    OperationResult SetSourceRange(IExcelBatch batch, string chartName, string sourceRange);

    /// <summary>
    /// Adds a data series to Regular Charts.
    /// PivotCharts: Returns error guiding to excel_pivottable(action: 'add-value-field').
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesName">Name for the series</param>
    /// <param name="valuesRange">Range containing Y values (e.g., "Sheet1!B2:B10")</param>
    /// <param name="categoryRange">Optional range for X values/categories</param>
    /// <returns>Series information</returns>
    ChartSeriesResult AddSeries(
        IExcelBatch batch,
        string chartName,
        string seriesName,
        string valuesRange,
        string? categoryRange = null);

    /// <summary>
    /// Removes a data series from Regular Charts.
    /// PivotCharts: Returns error guiding to excel_pivottable(action: 'remove-field').
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="seriesIndex">1-based index of series to remove</param>
    /// <returns>Operation result</returns>
    OperationResult RemoveSeries(IExcelBatch batch, string chartName, int seriesIndex);

    // === APPEARANCE OPERATIONS ===

    /// <summary>
    /// Changes chart type (works for both Regular and PivotCharts).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="chartType">New chart type from ChartType enum</param>
    /// <returns>Operation result</returns>
    OperationResult SetChartType(IExcelBatch batch, string chartName, ChartType chartType);

    /// <summary>
    /// Sets chart title.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="title">Chart title text (empty to hide title)</param>
    /// <returns>Operation result</returns>
    OperationResult SetTitle(IExcelBatch batch, string chartName, string title);

    /// <summary>
    /// Sets axis title.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="axis">Axis type (Primary, Secondary, Category, Value)</param>
    /// <param name="title">Axis title text</param>
    /// <returns>Operation result</returns>
    OperationResult SetAxisTitle(
        IExcelBatch batch,
        string chartName,
        ChartAxisType axis,
        string title);

    /// <summary>
    /// Shows or hides chart legend.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart</param>
    /// <param name="visible">True to show legend, false to hide</param>
    /// <param name="position">Legend position (optional)</param>
    /// <returns>Operation result</returns>
    OperationResult ShowLegend(
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
    /// <returns>Operation result</returns>
    OperationResult SetStyle(IExcelBatch batch, string chartName, int styleId);
}
