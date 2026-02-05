using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

/// <summary>
/// Excel chart lifecycle operations - creating, reading, moving, and deleting charts.
/// Supports two chart types: Regular (static, from ranges) and PivotCharts (dynamic, from PivotTables).
/// Configuration operations (series, titles, styling) are in IChartConfigCommands.
/// </summary>
[ServiceCategory("chart", "Chart")]
[McpTool("excel_chart")]
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
    [ServiceAction("delete")]
    void Delete(IExcelBatch batch, [RequiredParameter] string chartName);

    /// <summary>
    /// Moves/resizes a chart.
    /// </summary>
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
    [ServiceAction("fit-to-range")]
    void FitToRange(
        IExcelBatch batch,
        [RequiredParameter] string chartName,
        [RequiredParameter] string sheetName,
        [RequiredParameter] string rangeAddress);
}

