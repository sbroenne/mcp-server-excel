using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Chart lifecycle - create, read, move, and delete embedded charts.
///
/// CRITICAL - AVOID OVERLAPPING DATA:
/// 1. Check used range first with range get-used-range
/// 2. Position chart BELOW or to the RIGHT of data
/// 3. NEVER place charts at default position (0,0) - it overlaps data!
///
/// POSITIONING: left/top in points (72 points = 1 inch).
/// Use fit-to-range to position chart within a cell range like 'F2:K15'.
///
/// CHART TYPES: 70+ types available including Column, Line, Pie, Bar, Area, XY Scatter.
///
/// CREATE OPTIONS:
/// - create-from-range: Create from cell range (e.g., 'A1:D10')
/// - create-from-table: Create from Excel Table (uses table's data range)
/// - create-from-pivottable: Create linked PivotChart
///
/// Use chartconfig for series, titles, legends, styles, placement mode.
/// </summary>
[ServiceCategory("chart", "Chart")]
[McpTool("chart", Title = "Chart Operations", Destructive = true, Category = "analysis",
    Description = "Chart lifecycle - create, read, move, and delete embedded charts. CRITICAL - AVOID OVERLAPPING DATA: 1. Check used range first with get-used-range 2. Position chart BELOW or RIGHT of data 3. NEVER place at default (0,0). POSITIONING: targetRange (cell-relative, PREFERRED) or left/top (points, 72pts=1in). CHART TYPES: 70+ types (ColumnClustered, Line, Pie, Bar, Area, XYScatter, etc.). CREATE: create-from-range (cell range), create-from-table (Excel Table), create-from-pivottable (linked PivotChart). Use chart_config for series, titles, legends, and styling.")]
public interface IChartCommands
{
    // === LIFECYCLE OPERATIONS ===

    /// <summary>
    /// Lists all charts in workbook (Regular and PivotCharts).
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>List of charts with names, types, sheets, positions, data sources</returns>
    [ServiceAction("list")]
    List<ChartInfo> List(IExcelBatch batch);

    /// <summary>
    /// Gets complete chart configuration.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart (or shape name)</param>
    /// <returns>Chart type, data source, series info, position, styling</returns>
    [ServiceAction("read")]
    ChartInfoResult Read(IExcelBatch batch, [RequiredParameter] string chartName);

    /// <summary>
    /// Creates a Regular Chart from an Excel range.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="sourceRange">Data range for the chart (e.g., A1:D10)</param>
    /// <param name="chartType">Type of chart to create</param>
    /// <param name="left">Left position in points from worksheet edge</param>
    /// <param name="top">Top position in points from worksheet edge</param>
    /// <param name="width">Chart width in points</param>
    /// <param name="height">Chart height in points</param>
    /// <param name="chartName">Optional chart name (auto-generated if omitted)</param>
    [ServiceAction("create-from-range")]
    ChartCreateResult CreateFromRange(
        IExcelBatch batch,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string sourceRange,
        [RequiredParameter] ChartType chartType,
        double left,
        double top,
        double width = 400,
        double height = 300,
        string? chartName = null);

    /// <summary>
    /// Creates a Regular Chart from an Excel Table's data range.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="tableName">Name of the Excel Table</param>
    /// <param name="sheetName">Target worksheet name for the chart</param>
    /// <param name="chartType">Type of chart to create</param>
    /// <param name="left">Left position in points from worksheet edge</param>
    /// <param name="top">Top position in points from worksheet edge</param>
    /// <param name="width">Chart width in points</param>
    /// <param name="height">Chart height in points</param>
    /// <param name="chartName">Optional chart name (auto-generated if omitted)</param>
    [ServiceAction("create-from-table")]
    ChartCreateResult CreateFromTable(
        IExcelBatch batch,
        [RequiredParameter] string tableName,
        [RequiredParameter] string sheetName,
        [RequiredParameter] ChartType chartType,
        double left,
        double top,
        double width = 400,
        double height = 300,
        string? chartName = null);

    /// <summary>
    /// Creates a PivotChart from an existing PivotTable.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the source PivotTable</param>
    /// <param name="sheetName">Target worksheet name for the chart</param>
    /// <param name="chartType">Type of chart to create</param>
    /// <param name="left">Left position in points from worksheet edge</param>
    /// <param name="top">Top position in points from worksheet edge</param>
    /// <param name="width">Chart width in points</param>
    /// <param name="height">Chart height in points</param>
    /// <param name="chartName">Optional chart name (auto-generated if omitted)</param>
    [ServiceAction("create-from-pivottable")]
    ChartCreateResult CreateFromPivotTable(
        IExcelBatch batch,
        [RequiredParameter] string pivotTableName,
        [RequiredParameter] string sheetName,
        [RequiredParameter] ChartType chartType,
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
    [ServiceAction("delete")]
    void Delete(IExcelBatch batch, [RequiredParameter] string chartName);

    /// <summary>
    /// Moves/resizes a chart.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart to move</param>
    /// <param name="left">New left position in points (null to keep current)</param>
    /// <param name="top">New top position in points (null to keep current)</param>
    /// <param name="width">New width in points (null to keep current)</param>
    /// <param name="height">New height in points (null to keep current)</param>
    [ServiceAction("move")]
    void Move(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        double? left = null,
        double? top = null,
        double? width = null,
        double? height = null);

    /// <summary>
    /// Fits a chart to a cell range by setting position and size to match the range bounds.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="chartName">Name of the chart to fit</param>
    /// <param name="sheetName">Worksheet containing the range</param>
    /// <param name="rangeAddress">Range to fit the chart to (e.g., A1:D10)</param>
    [ServiceAction("fit-to-range")]
    void FitToRange(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress);
}

