using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Chart lifecycle operations.
/// Use excel_chart_config for series, appearance, and styling.
/// </summary>
[McpServerToolType]
public static partial class ExcelChartTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// Chart lifecycle - create, read, move, and delete embedded charts.
    ///
    /// CHART TYPES: 70+ types available including:
    /// - Column: ColumnClustered, ColumnStacked, ColumnStacked100, Column3D
    /// - Line: Line, LineMarkers, LineStacked, LineMarkerStacked
    /// - Pie: Pie, Pie3D, PieExploded, Doughnut
    /// - Bar: BarClustered, BarStacked, BarStacked100
    /// - Area: Area, AreaStacked, AreaStacked100
    /// - XY Scatter: XYScatter, XYScatterLines, XYScatterSmooth
    ///
    /// POSITIONING: Use left/top in points (1 inch = 72 points).
    /// Default size is 400x300 points.
    ///
    /// PIVOTCHART: Use create-from-pivottable to link chart to PivotTable.
    /// Changes to PivotTable filters automatically update the chart.
    ///
    /// RELATED TOOLS:
    /// - excel_chart_config: Series, titles, legends, styles
    /// - excel_pivottable: Create PivotTables for PivotCharts
    /// </summary>
    /// <param name="action">The chart operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="chartName">Name of the chart - optional for create (auto-generated), required for other actions</param>
    /// <param name="sheetName">Worksheet name where chart will be placed (required for create actions)</param>
    /// <param name="sourceRange">Data range for chart like 'A1:D10' (required for create-from-range)</param>
    /// <param name="chartType">Chart type enum like ColumnClustered, Line, Pie</param>
    /// <param name="pivotTableName">PivotTable name (required for create-from-pivottable action)</param>
    /// <param name="left">Left position in points (72 points = 1 inch)</param>
    /// <param name="top">Top position in points (72 points = 1 inch)</param>
    /// <param name="width">Chart width in points (default 400)</param>
    /// <param name="height">Chart height in points (default 300)</param>
    [McpServerTool(Name = "excel_chart", Title = "Excel Chart Operations")]
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
        [DefaultValue(null)] double? left,
        [DefaultValue(null)] double? top,
        [DefaultValue(null)] double? width,
        [DefaultValue(null)] double? height)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_chart",
            action.ToActionString(),
            () =>
            {
                var commands = new ChartCommands();

                return action switch
                {
                    ChartAction.List => ListAction(commands, sessionId),
                    ChartAction.Read => ReadAction(commands, sessionId, chartName),
                    ChartAction.CreateFromRange => CreateFromRangeAction(commands, sessionId, sheetName, sourceRange, chartType, left, top, width, height, chartName),
                    ChartAction.CreateFromPivotTable => CreateFromPivotTableAction(commands, sessionId, pivotTableName, sheetName, chartType, left, top, width, height, chartName),
                    ChartAction.Delete => DeleteAction(commands, sessionId, chartName),
                    ChartAction.Move => MoveAction(commands, sessionId, chartName, left, top, width, height),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListAction(
        ChartCommands commands,
        string sessionId)
    {
        var charts = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));
        return JsonSerializer.Serialize(charts, JsonOptions);
    }

    private static string ReadAction(
        ChartCommands commands,
        string sessionId,
        string? chartName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "read");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Read(batch, chartName!));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string CreateFromRangeAction(
        ChartCommands commands,
        string sessionId,
        string? sheetName,
        string? sourceRange,
        ChartType? chartType,
        double? left,
        double? top,
        double? width,
        double? height,
        string? chartName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create-from-range");
        if (string.IsNullOrWhiteSpace(sourceRange))
            ExcelToolsBase.ThrowMissingParameter(nameof(sourceRange), "create-from-range");
        if (!chartType.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(chartType), "create-from-range");
        if (!left.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(left), "create-from-range");
        if (!top.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(top), "create-from-range");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateFromRange(batch, sheetName!, sourceRange!, chartType!.Value,
                left!.Value, top!.Value, width ?? 400, height ?? 300, chartName));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string CreateFromPivotTableAction(
        ChartCommands commands,
        string sessionId,
        string? pivotTableName,
        string? sheetName,
        ChartType? chartType,
        double? left,
        double? top,
        double? width,
        double? height,
        string? chartName)
    {
        if (string.IsNullOrWhiteSpace(pivotTableName))
            ExcelToolsBase.ThrowMissingParameter(nameof(pivotTableName), "create-from-pivottable");
        if (string.IsNullOrWhiteSpace(sheetName))
            ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create-from-pivottable");
        if (!chartType.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(chartType), "create-from-pivottable");
        if (!left.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(left), "create-from-pivottable");
        if (!top.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(top), "create-from-pivottable");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.CreateFromPivotTable(batch, pivotTableName!, sheetName!, chartType!.Value,
                left!.Value, top!.Value, width ?? 400, height ?? 300, chartName));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string DeleteAction(
        ChartCommands commands,
        string sessionId,
        string? chartName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "delete");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.Delete(batch, chartName!);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' deleted successfully" }, JsonOptions);
    }

    private static string MoveAction(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        double? left,
        double? top,
        double? width,
        double? height)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "move");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.Move(batch, chartName!, left, top, width, height);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' moved successfully" }, JsonOptions);
    }


}
