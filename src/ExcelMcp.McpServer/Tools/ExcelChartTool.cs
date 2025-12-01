using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Chart operations
/// </summary>
[McpServerToolType]
public static partial class ExcelChartTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// Excel Chart operations - create, modify, and analyze Regular Charts and PivotCharts.
    /// CHART TYPES: Supports 70+ chart types including Column, Line, Pie, Bar, Area, XYScatter, and specialized types.
    /// STRATEGY PATTERN: Automatically handles differences between Regular Charts (use SeriesCollection) and PivotCharts (read-only data source).
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Path to Excel file (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="chartName">Chart name (required for most operations)</param>
    /// <param name="sheetName">Sheet name (for create-from-range)</param>
    /// <param name="sourceRange">Source data range (e.g., 'A1:D10' for create-from-range, set-source-range)</param>
    /// <param name="chartType">Chart type enum value (e.g., ColumnClustered, Line, Pie, BarClustered, Area, XYScatter)</param>
    /// <param name="pivotTableName">PivotTable name (for create-from-pivottable)</param>
    /// <param name="left">Left position in points</param>
    /// <param name="top">Top position in points</param>
    /// <param name="width">Width in points (optional, uses Excel default if not specified)</param>
    /// <param name="height">Height in points (optional, uses Excel default if not specified)</param>
    /// <param name="title">Chart title text (empty string hides title)</param>
    /// <param name="axis">Axis type for set-axis-title: Category, Value, CategorySecondary, ValueSecondary</param>
    /// <param name="seriesName">Series name (for add-series)</param>
    /// <param name="valuesRange">Values range for series (e.g., 'Sheet1!B2:B10' for add-series)</param>
    /// <param name="categoryRange">Category range for series (e.g., 'Sheet1!A2:A10', optional for add-series)</param>
    /// <param name="seriesIndex">1-based series index (for remove-series)</param>
    /// <param name="visible">Show/hide legend: true=show, false=hide</param>
    /// <param name="legendPosition">Legend position: Bottom, Corner, Top, Right, Left</param>
    /// <param name="styleId">Chart style ID (1-48)</param>
    [McpServerTool(Name = "excel_chart")]
    public static partial string ExcelChart(
        ChartAction action,
        string excelPath,
        string sessionId,
        string? chartName,
        string? sheetName,
        string? sourceRange,
        ChartType? chartType,
        string? pivotTableName,
        double? left,
        double? top,
        double? width,
        double? height,
        string? title,
        ChartAxisType? axis,
        string? seriesName,
        string? valuesRange,
        string? categoryRange,
        int? seriesIndex,
        bool? visible,
        LegendPosition? legendPosition,
        int? styleId)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_chart",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var commands = new ChartCommands();

                return action switch
                {
                    ChartAction.List => List(commands, sessionId),
                    ChartAction.Read => Read(commands, sessionId, chartName),
                    ChartAction.CreateFromRange => CreateFromRange(commands, sessionId, sheetName, sourceRange, chartType, left, top, width, height, chartName),
                    ChartAction.CreateFromPivotTable => CreateFromPivotTable(commands, sessionId, pivotTableName, sheetName, chartType, left, top, width, height, chartName),
                    ChartAction.Delete => Delete(commands, sessionId, chartName),
                    ChartAction.Move => Move(commands, sessionId, chartName, left, top, width, height),
                    ChartAction.SetSourceRange => SetSourceRange(commands, sessionId, chartName, sourceRange),
                    ChartAction.AddSeries => AddSeries(commands, sessionId, chartName, seriesName, valuesRange, categoryRange),
                    ChartAction.RemoveSeries => RemoveSeries(commands, sessionId, chartName, seriesIndex),
                    ChartAction.SetChartType => SetChartType(commands, sessionId, chartName, chartType),
                    ChartAction.SetTitle => SetTitle(commands, sessionId, chartName, title),
                    ChartAction.SetAxisTitle => SetAxisTitle(commands, sessionId, chartName, axis, title),
                    ChartAction.ShowLegend => ShowLegend(commands, sessionId, chartName, visible, legendPosition),
                    ChartAction.SetStyle => SetStyle(commands, sessionId, chartName, styleId),
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

        return JsonSerializer.Serialize(new { success = true, message = $"Chart '{chartName}' deleted successfully" }, JsonOptions);
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

        return JsonSerializer.Serialize(new { success = true, message = $"Chart '{chartName}' moved successfully" }, JsonOptions);
    }

    private static string SetSourceRange(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        string? sourceRange)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-source-range");
        if (string.IsNullOrWhiteSpace(sourceRange))
            ExcelToolsBase.ThrowMissingParameter(nameof(sourceRange), "set-source-range");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetSourceRange(batch, chartName!, sourceRange!);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new { success = true, message = $"Chart '{chartName}' source range set to '{sourceRange}'" }, JsonOptions);
    }

    private static string AddSeries(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        string? seriesName,
        string? valuesRange,
        string? categoryRange)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "add-series");
        if (string.IsNullOrWhiteSpace(seriesName))
            ExcelToolsBase.ThrowMissingParameter(nameof(seriesName), "add-series");
        if (string.IsNullOrWhiteSpace(valuesRange))
            ExcelToolsBase.ThrowMissingParameter(nameof(valuesRange), "add-series");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddSeries(batch, chartName!, seriesName!, valuesRange!, categoryRange));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string RemoveSeries(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        int? seriesIndex)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "remove-series");
        if (!seriesIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(seriesIndex), "remove-series");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.RemoveSeries(batch, chartName!, seriesIndex!.Value);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new { success = true, message = $"Series {seriesIndex!.Value} removed from chart '{chartName}'" }, JsonOptions);
    }

    private static string SetChartType(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        ChartType? chartType)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-chart-type");
        if (!chartType.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(chartType), "set-chart-type");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetChartType(batch, chartName!, chartType!.Value);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new { success = true, message = $"Chart '{chartName}' type changed to {chartType!.Value}" }, JsonOptions);
    }

    private static string SetTitle(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        string? title)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-title");
        if (title == null)
            ExcelToolsBase.ThrowMissingParameter(nameof(title), "set-title");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetTitle(batch, chartName!, title!);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new { success = true, message = string.IsNullOrEmpty(title) ? $"Chart '{chartName}' title hidden" : $"Chart '{chartName}' title set" }, JsonOptions);
    }

    private static string SetAxisTitle(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        ChartAxisType? axis,
        string? title)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-axis-title");
        if (!axis.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(axis), "set-axis-title");
        if (title == null)
            ExcelToolsBase.ThrowMissingParameter(nameof(title), "set-axis-title");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetAxisTitle(batch, chartName!, axis!.Value, title!);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new { success = true, message = $"Chart '{chartName}' {axis!.Value} axis title set" }, JsonOptions);
    }

    private static string ShowLegend(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        bool? visible,
        LegendPosition? legendPosition)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "show-legend");
        if (!visible.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(visible), "show-legend");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.ShowLegend(batch, chartName!, visible!.Value, legendPosition);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new { success = true, message = visible!.Value ? $"Chart '{chartName}' legend shown" : $"Chart '{chartName}' legend hidden" }, JsonOptions);
    }

    private static string SetStyle(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        int? styleId)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-style");
        if (!styleId.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(styleId), "set-style");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetStyle(batch, chartName!, styleId!.Value);
            return 0; // Dummy return value
        });

        return JsonSerializer.Serialize(new { success = true, message = $"Chart '{chartName}' style set to {styleId!.Value}" }, JsonOptions);
    }
}
