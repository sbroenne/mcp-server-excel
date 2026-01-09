using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Chart configuration operations
/// </summary>
[McpServerToolType]
public static partial class ExcelChartConfigTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// Chart configuration - data source, series, type, title, axis, legend, style.
    /// CHART TYPES: 70+ (Column, Line, Pie, Bar, Area, XYScatter, etc.)
    /// Related: excel_chart (lifecycle)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="sid">Session ID</param>
    /// <param name="cn">Chart name</param>
    /// <param name="sr">Source range (e.g., 'A1:D10')</param>
    /// <param name="ct">Chart type enum</param>
    /// <param name="title">Title text (empty string hides)</param>
    /// <param name="axis">Axis: Category, Value, CategorySecondary, ValueSecondary</param>
    /// <param name="sn">Series name</param>
    /// <param name="vr">Values range (e.g., 'Sheet1!B2:B10')</param>
    /// <param name="cr">Category range (optional)</param>
    /// <param name="si">Series index (1-based)</param>
    /// <param name="vis">Show legend (true/false)</param>
    /// <param name="lp">Legend position: Bottom, Corner, Top, Right, Left</param>
    /// <param name="style">Chart style ID (1-48)</param>
    [McpServerTool(Name = "excel_chart_config", Title = "Excel Chart Configuration")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelChartConfig(
        ChartConfigAction action,
        string sid,
        string cn,
        [DefaultValue(null)] string? sr,
        [DefaultValue(null)] ChartType? ct,
        [DefaultValue(null)] string? title,
        [DefaultValue(null)] ChartAxisType? axis,
        [DefaultValue(null)] string? sn,
        [DefaultValue(null)] string? vr,
        [DefaultValue(null)] string? cr,
        [DefaultValue(null)] int? si,
        [DefaultValue(null)] bool? vis,
        [DefaultValue(null)] LegendPosition? lp,
        [DefaultValue(null)] int? style)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_chart_config",
            action.ToActionString(),
            () =>
            {
                var commands = new ChartCommands();

                return action switch
                {
                    ChartConfigAction.SetSourceRange => SetSourceRange(commands, sid, cn, sr),
                    ChartConfigAction.AddSeries => AddSeries(commands, sid, cn, sn, vr, cr),
                    ChartConfigAction.RemoveSeries => RemoveSeries(commands, sid, cn, si),
                    ChartConfigAction.SetChartType => SetChartType(commands, sid, cn, ct),
                    ChartConfigAction.SetTitle => SetTitle(commands, sid, cn, title),
                    ChartConfigAction.SetAxisTitle => SetAxisTitle(commands, sid, cn, axis, title),
                    ChartConfigAction.ShowLegend => ShowLegend(commands, sid, cn, vis, lp),
                    ChartConfigAction.SetStyle => SetStyle(commands, sid, cn, style),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string SetSourceRange(ChartCommands commands, string sessionId, string? chartName, string? sourceRange)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter("cn", "set-source-range");
        if (string.IsNullOrWhiteSpace(sourceRange))
            ExcelToolsBase.ThrowMissingParameter("sr", "set-source-range");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetSourceRange(batch, chartName!, sourceRange!);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' source range set to '{sourceRange}'" }, JsonOptions);
    }

    private static string AddSeries(ChartCommands commands, string sessionId, string? chartName, string? seriesName, string? valuesRange, string? categoryRange)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter("cn", "add-series");
        if (string.IsNullOrWhiteSpace(seriesName))
            ExcelToolsBase.ThrowMissingParameter("sn", "add-series");
        if (string.IsNullOrWhiteSpace(valuesRange))
            ExcelToolsBase.ThrowMissingParameter("vr", "add-series");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddSeries(batch, chartName!, seriesName!, valuesRange!, categoryRange));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string RemoveSeries(ChartCommands commands, string sessionId, string? chartName, int? seriesIndex)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter("cn", "remove-series");
        if (!seriesIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter("si", "remove-series");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.RemoveSeries(batch, chartName!, seriesIndex!.Value);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Series {seriesIndex!.Value} removed from chart '{chartName}'" }, JsonOptions);
    }

    private static string SetChartType(ChartCommands commands, string sessionId, string? chartName, ChartType? chartType)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter("cn", "set-chart-type");
        if (!chartType.HasValue)
            ExcelToolsBase.ThrowMissingParameter("ct", "set-chart-type");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetChartType(batch, chartName!, chartType!.Value);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' type changed to {chartType!.Value}" }, JsonOptions);
    }

    private static string SetTitle(ChartCommands commands, string sessionId, string? chartName, string? title)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter("cn", "set-title");
        if (title == null)
            ExcelToolsBase.ThrowMissingParameter("title", "set-title");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetTitle(batch, chartName!, title!);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = string.IsNullOrEmpty(title) ? $"Chart '{chartName}' title hidden" : $"Chart '{chartName}' title set" }, JsonOptions);
    }

    private static string SetAxisTitle(ChartCommands commands, string sessionId, string? chartName, ChartAxisType? axis, string? title)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter("cn", "set-axis-title");
        if (!axis.HasValue)
            ExcelToolsBase.ThrowMissingParameter("axis", "set-axis-title");
        if (title == null)
            ExcelToolsBase.ThrowMissingParameter("title", "set-axis-title");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetAxisTitle(batch, chartName!, axis!.Value, title!);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' {axis!.Value} axis title set" }, JsonOptions);
    }

    private static string ShowLegend(ChartCommands commands, string sessionId, string? chartName, bool? visible, LegendPosition? legendPosition)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter("cn", "show-legend");
        if (!visible.HasValue)
            ExcelToolsBase.ThrowMissingParameter("vis", "show-legend");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.ShowLegend(batch, chartName!, visible!.Value, legendPosition);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = visible!.Value ? $"Chart '{chartName}' legend shown" : $"Chart '{chartName}' legend hidden" }, JsonOptions);
    }

    private static string SetStyle(ChartCommands commands, string sessionId, string? chartName, int? styleId)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter("cn", "set-style");
        if (!styleId.HasValue)
            ExcelToolsBase.ThrowMissingParameter("style", "set-style");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetStyle(batch, chartName!, styleId!.Value);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' style set to {styleId!.Value}" }, JsonOptions);
    }
}
