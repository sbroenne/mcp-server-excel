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
    /// Chart lifecycle. CHART TYPES: 70+ (Column, Line, Pie, Bar, Area, XYScatter, etc.)
    /// Related: excel_chart_config (series/titles/legends/styles)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="sid">Session ID</param>
    /// <param name="cn">Chart name</param>
    /// <param name="sn">Sheet name</param>
    /// <param name="sr">Source data range (A1:D10)</param>
    /// <param name="ct">Chart type (ColumnClustered, Line, Pie, etc.)</param>
    /// <param name="ptn">PivotTable name (create-from-pivottable)</param>
    /// <param name="left">Left position (points)</param>
    /// <param name="top">Top position (points)</param>
    /// <param name="width">Width (points, default 400)</param>
    /// <param name="height">Height (points, default 300)</param>
    [McpServerTool(Name = "excel_chart", Title = "Excel Chart Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelChart(
        ChartAction action,
        string sid,
        [DefaultValue(null)] string? cn,
        [DefaultValue(null)] string? sn,
        [DefaultValue(null)] string? sr,
        [DefaultValue(null)] ChartType? ct,
        [DefaultValue(null)] string? ptn,
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
                    ChartAction.List => List(commands, sid),
                    ChartAction.Read => Read(commands, sid, cn),
                    ChartAction.CreateFromRange => CreateFromRange(commands, sid, sn, sr, ct, left, top, width, height, cn),
                    ChartAction.CreateFromPivotTable => CreateFromPivotTable(commands, sid, ptn, sn, ct, left, top, width, height, cn),
                    ChartAction.Delete => Delete(commands, sid, cn),
                    ChartAction.Move => Move(commands, sid, cn, left, top, width, height),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string List(
        ChartCommands commands,
        string sessionId)
    {
        var charts = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));
        return JsonSerializer.Serialize(charts, JsonOptions);
    }

    private static string Read(
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

    private static string CreateFromRange(
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

    private static string CreateFromPivotTable(
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

    private static string Delete(
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

    private static string Move(
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
