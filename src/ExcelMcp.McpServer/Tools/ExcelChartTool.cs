using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Chart lifecycle operations.
/// Use excel_chart_config for series, appearance, and styling.
/// </summary>
[McpServerToolType]
public static partial class ExcelChartTool
{
    /// <summary>
    /// Chart lifecycle - create, read, move, and delete embedded charts.
    ///
    /// CRITICAL - AVOID OVERLAPPING DATA:
    /// 1. Check used range first: excel_range(action: 'get-used-range')
    /// 2. Position chart BELOW or to the RIGHT of data
    /// 3. Use targetRange for cell-relative positioning (RECOMMENDED)
    /// 4. NEVER place charts at default position (0,0) - it overlaps data!
    ///
    /// POSITIONING OPTIONS (for create actions):
    /// - targetRange: Position by cell range (e.g., 'F2:K15') - PREFERRED
    /// - left/top: Position by points (72 points = 1 inch)
    /// One of these is required for create-from-range and create-from-pivottable.
    ///
    /// CHART TYPES: 70+ types available including:
    /// - Column: ColumnClustered, ColumnStacked, ColumnStacked100, Column3D
    /// - Line: Line, LineMarkers, LineStacked, LineMarkerStacked
    /// - Pie: Pie, Pie3D, PieExploded, Doughnut
    /// - Bar: BarClustered, BarStacked, BarStacked100
    /// - Area: Area, AreaStacked, AreaStacked100
    /// - XY Scatter: XYScatter, XYScatterLines, XYScatterSmooth
    ///
    /// CREATE OPTIONS:
    /// - create-from-range: Create from cell range (e.g., 'A1:D10')
    /// - create-from-table: Create from Excel Table (uses table's data range)
    /// - create-from-pivottable: Create linked PivotChart
    ///
    /// PIVOTCHART: Use create-from-pivottable to link chart to PivotTable.
    /// Changes to PivotTable filters automatically update the chart.
    ///
    /// RELATED TOOLS:
    /// - excel_chart_config: Series, titles, legends, styles, placement mode
    /// - excel_pivottable: Create PivotTables for PivotCharts
    /// - excel_range: Get cell geometry for positioning
    /// </summary>
    /// <param name="action">The chart operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="chartName">Name of the chart - optional for create (auto-generated), required for other actions</param>
    /// <param name="sheetName">Worksheet name where chart will be placed (required for create and fit-to-range actions)</param>
    /// <param name="sourceRange">Data range for chart like 'A1:D10' (required for create-from-range)</param>
    /// <param name="chartType">Chart type enum like ColumnClustered, Line, Pie</param>
    /// <param name="pivotTableName">PivotTable name (required for create-from-pivottable action)</param>
    /// <param name="tableName">Excel Table name (required for create-from-table action)</param>
    /// <param name="left">Left position in points (72 points = 1 inch)</param>
    /// <param name="top">Top position in points (72 points = 1 inch)</param>
    /// <param name="width">Chart width in points (default 400)</param>
    /// <param name="height">Chart height in points (default 300)</param>
    /// <param name="rangeAddress">Target cell range for fit-to-range action (e.g., 'A1:D10')</param>
    /// <param name="targetRange">Position chart within this cell range (e.g., 'F2:K15') - alternative to left/top for create actions</param>
    [McpServerTool(Name = "excel_chart", Title = "Excel Chart Operations", Destructive = true)]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelChart(
        ChartAction action,
        string sessionId,
        [DefaultValue(null)] string? chartName,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? sourceRange,
        [DefaultValue(null)] ChartType? chartType,
        [DefaultValue(null)] string? pivotTableName,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] double? left,
        [DefaultValue(null)] double? top,
        [DefaultValue(null)] double? width,
        [DefaultValue(null)] double? height,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] string? targetRange)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_chart",
            ServiceRegistry.Chart.ToActionString(action),
            () =>
            {
                return action switch
                {
                    ChartAction.List => ExcelToolsBase.ForwardToService("chart.list", sessionId),
                    ChartAction.Read => ForwardRead(sessionId, chartName),
                    ChartAction.CreateFromRange => ForwardCreateFromRange(sessionId, sheetName, sourceRange, chartType, left, top, width, height, chartName, targetRange),
                    ChartAction.CreateFromTable => ForwardCreateFromTable(sessionId, tableName, sheetName, chartType, left, top, width, height, chartName, targetRange),
                    ChartAction.CreateFromPivotTable => ForwardCreateFromPivotTable(sessionId, pivotTableName, sheetName, chartType, left, top, width, height, chartName, targetRange),
                    ChartAction.Delete => ForwardDelete(sessionId, chartName),
                    ChartAction.Move => ForwardMove(sessionId, chartName, left, top, width, height),
                    ChartAction.FitToRange => ForwardFitToRange(sessionId, chartName, sheetName, rangeAddress),
                    _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.Chart.ToActionString(action)})", nameof(action))
                };
            });
    }

    // === SERVICE FORWARDING METHODS ===

    private static string ForwardRead(string sessionId, string? chartName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "read");

        return ExcelToolsBase.ForwardToService("chart.read", sessionId, new { chartName });
    }

    private static string ForwardCreateFromRange(
        string sessionId,
        string? sheetName,
        string? sourceRange,
        ChartType? chartType,
        double? left,
        double? top,
        double? width,
        double? height,
        string? chartName,
        string? targetRange)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create-from-range");
        if (string.IsNullOrWhiteSpace(sourceRange))
            ExcelToolsBase.ThrowMissingParameter(nameof(sourceRange), "create-from-range");
        if (!chartType.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(chartType), "create-from-range");

        // Either targetRange OR (left, top) must be provided
        bool hasTargetRange = !string.IsNullOrWhiteSpace(targetRange);
        bool hasPointPosition = left.HasValue && top.HasValue;

        if (!hasTargetRange && !hasPointPosition)
            ExcelToolsBase.ThrowMissingParameter("targetRange or left/top", "create-from-range");

        // Convert ChartType enum to string for service
        var chartTypeString = chartType!.Value.ToString();

        return ExcelToolsBase.ForwardToService("chart.create-from-range", sessionId, new
        {
            sheetName,
            sourceRange,
            chartType = chartTypeString,
            left = hasTargetRange ? 0 : left,
            top = hasTargetRange ? 0 : top,
            width = width ?? 400,
            height = height ?? 300,
            chartName,
            targetRange
        });
    }

    private static string ForwardCreateFromTable(
        string sessionId,
        string? tableName,
        string? sheetName,
        ChartType? chartType,
        double? left,
        double? top,
        double? width,
        double? height,
        string? chartName,
        string? targetRange)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create-from-table");
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create-from-table");
        if (!chartType.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(chartType), "create-from-table");

        // Either targetRange OR (left, top) must be provided
        bool hasTargetRange = !string.IsNullOrWhiteSpace(targetRange);
        bool hasPointPosition = left.HasValue && top.HasValue;

        if (!hasTargetRange && !hasPointPosition)
            ExcelToolsBase.ThrowMissingParameter("targetRange or left/top", "create-from-table");

        var chartTypeString = chartType!.Value.ToString();

        return ExcelToolsBase.ForwardToService("chart.create-from-table", sessionId, new
        {
            tableName,
            sheetName,
            chartType = chartTypeString,
            left = hasTargetRange ? 0 : left,
            top = hasTargetRange ? 0 : top,
            width = width ?? 400,
            height = height ?? 300,
            chartName,
            targetRange
        });
    }

    private static string ForwardCreateFromPivotTable(
        string sessionId,
        string? pivotTableName,
        string? sheetName,
        ChartType? chartType,
        double? left,
        double? top,
        double? width,
        double? height,
        string? chartName,
        string? targetRange)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-pivottable");
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create-from-pivottable");
        if (!chartType.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(chartType), "create-from-pivottable");

        // Either targetRange OR (left, top) must be provided
        bool hasTargetRange = !string.IsNullOrWhiteSpace(targetRange);
        bool hasPointPosition = left.HasValue && top.HasValue;

        if (!hasTargetRange && !hasPointPosition)
            ExcelToolsBase.ThrowMissingParameter("targetRange or left/top", "create-from-pivottable");

        var chartTypeString = chartType!.Value.ToString();

        return ExcelToolsBase.ForwardToService("chart.create-from-pivottable", sessionId, new
        {
            pivotTableName,
            sheetName,
            chartType = chartTypeString,
            left = hasTargetRange ? 0 : left,
            top = hasTargetRange ? 0 : top,
            width = width ?? 400,
            height = height ?? 300,
            chartName,
            targetRange
        });
    }

    private static string ForwardDelete(string sessionId, string? chartName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "delete");

        return ExcelToolsBase.ForwardToService("chart.delete", sessionId, new { chartName });
    }

    private static string ForwardMove(
        string sessionId,
        string? chartName,
        double? left,
        double? top,
        double? width,
        double? height)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "move");

        return ExcelToolsBase.ForwardToService("chart.move", sessionId, new { chartName, left, top, width, height });
    }

    private static string ForwardFitToRange(
        string sessionId,
        string? chartName,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "fit-to-range");
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "fit-to-range");
        if (string.IsNullOrWhiteSpace(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter(nameof(rangeAddress), "fit-to-range");

        return ExcelToolsBase.ForwardToService("chart.fit-to-range", sessionId, new { chartName, sheetName, rangeAddress });
    }
}




