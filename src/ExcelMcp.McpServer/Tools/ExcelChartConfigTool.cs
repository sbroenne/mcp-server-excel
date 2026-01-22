using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Chart configuration operations.
/// Use excel_chart for chart lifecycle (create, delete, move).
/// </summary>
[McpServerToolType]
public static partial class ExcelChartConfigTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    /// <summary>
    /// Chart configuration - data source, series, type, title, axis labels, legend, and styling.
    ///
    /// CHART TYPES: 70+ types available (ColumnClustered, Line, Pie, Bar, Area, XYScatter, etc.)
    ///
    /// SERIES MANAGEMENT:
    /// - add-series: Add a new data series with valuesRange (required) and optional categoryRange
    /// - remove-series: Remove series by 1-based index
    /// - set-source-range: Replace entire chart data source
    ///
    /// TITLES AND LABELS:
    /// - set-title: Set chart title (empty string hides title)
    /// - set-axis-title: Set axis labels (Category, Value, CategorySecondary, ValueSecondary)
    ///
    /// AXIS FORMATTING:
    /// - get-axis-number-format: Get current number format for axis tick labels
    /// - set-axis-number-format: Set number format for axis tick labels (e.g., "$#,##0,,"M"" for millions)
    ///
    /// LEGEND POSITIONS: Bottom, Corner, Top, Right, Left
    ///
    /// CHART STYLES: 1-48 (built-in Excel styles with different color schemes)
    ///
    /// RELATED TOOLS:
    /// - excel_chart: Create, delete, and move charts
    /// </summary>
    /// <param name="action">The chart configuration action to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="chartName">Name of the chart to configure (required for all actions)</param>
    /// <param name="sourceRange">Data range like 'A1:D10' or 'Sheet1!A1:D10' for set-source-range</param>
    /// <param name="chartType">Chart type enum for set-chart-type action</param>
    /// <param name="title">Title text for set-title and set-axis-title (empty string hides)</param>
    /// <param name="axis">Axis type: Category, Value, CategorySecondary, ValueSecondary</param>
    /// <param name="numberFormat">Excel number format string for set-axis-number-format (e.g., "$#,##0,,"M"" for millions, "0%" for percentage)</param>
    /// <param name="seriesName">Name for new series in add-series action</param>
    /// <param name="valuesRange">Data range for series values like 'Sheet1!B2:B10' (required for add-series)</param>
    /// <param name="categoryRange">Optional category range for series X-axis labels</param>
    /// <param name="seriesIndex">1-based series index for remove-series action</param>
    /// <param name="visible">Show or hide legend in show-legend action</param>
    /// <param name="legendPosition">Legend position: Bottom, Corner, Top, Right, Left</param>
    /// <param name="styleId">Chart style ID from 1-48 for set-style action</param>
    [McpServerTool(Name = "excel_chart_config", Title = "Excel Chart Configuration", Destructive = true)]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelChartConfig(
        ChartConfigAction action,
        string sessionId,
        string chartName,
        [DefaultValue(null)] string? sourceRange,
        [DefaultValue(null)] ChartType? chartType,
        [DefaultValue(null)] string? title,
        [DefaultValue(null)] ChartAxisType? axis,
        [DefaultValue(null)] string? numberFormat,
        [DefaultValue(null)] string? seriesName,
        [DefaultValue(null)] string? valuesRange,
        [DefaultValue(null)] string? categoryRange,
        [DefaultValue(null)] int? seriesIndex,
        [DefaultValue(null)] bool? visible,
        [DefaultValue(null)] LegendPosition? legendPosition,
        [DefaultValue(null)] int? styleId)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_chart_config",
            action.ToActionString(),
            () =>
            {
                var commands = new ChartCommands();

                return action switch
                {
                    ChartConfigAction.SetSourceRange => SetSourceRangeAction(commands, sessionId, chartName, sourceRange),
                    ChartConfigAction.AddSeries => AddSeriesAction(commands, sessionId, chartName, seriesName, valuesRange, categoryRange),
                    ChartConfigAction.RemoveSeries => RemoveSeriesAction(commands, sessionId, chartName, seriesIndex),
                    ChartConfigAction.SetChartType => SetChartTypeAction(commands, sessionId, chartName, chartType),
                    ChartConfigAction.SetTitle => SetTitleAction(commands, sessionId, chartName, title),
                    ChartConfigAction.SetAxisTitle => SetAxisTitleAction(commands, sessionId, chartName, axis, title),
                    ChartConfigAction.GetAxisNumberFormat => GetAxisNumberFormatAction(commands, sessionId, chartName, axis),
                    ChartConfigAction.SetAxisNumberFormat => SetAxisNumberFormatAction(commands, sessionId, chartName, axis, numberFormat),
                    ChartConfigAction.ShowLegend => ShowLegendAction(commands, sessionId, chartName, visible, legendPosition),
                    ChartConfigAction.SetStyle => SetStyleAction(commands, sessionId, chartName, styleId),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string SetSourceRangeAction(ChartCommands commands, string sessionId, string? chartName, string? sourceRange)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-source-range");
        if (string.IsNullOrWhiteSpace(sourceRange))
            ExcelToolsBase.ThrowMissingParameter(nameof(sourceRange), "set-source-range");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetSourceRange(batch, chartName!, sourceRange!);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' source range set to '{sourceRange}'" }, JsonOptions);
    }

    private static string AddSeriesAction(ChartCommands commands, string sessionId, string? chartName, string? seriesName, string? valuesRange, string? categoryRange)
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

    private static string RemoveSeriesAction(ChartCommands commands, string sessionId, string? chartName, int? seriesIndex)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "remove-series");
        if (!seriesIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(seriesIndex), "remove-series");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.RemoveSeries(batch, chartName!, seriesIndex!.Value);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Series {seriesIndex!.Value} removed from chart '{chartName}'" }, JsonOptions);
    }

    private static string SetChartTypeAction(ChartCommands commands, string sessionId, string? chartName, ChartType? chartType)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-chart-type");
        if (!chartType.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(chartType), "set-chart-type");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetChartType(batch, chartName!, chartType!.Value);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' type changed to {chartType!.Value}" }, JsonOptions);
    }

    private static string SetTitleAction(ChartCommands commands, string sessionId, string? chartName, string? title)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-title");
        if (title == null)
            ExcelToolsBase.ThrowMissingParameter(nameof(title), "set-title");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetTitle(batch, chartName!, title!);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = string.IsNullOrEmpty(title) ? $"Chart '{chartName}' title hidden" : $"Chart '{chartName}' title set" }, JsonOptions);
    }

    private static string SetAxisTitleAction(ChartCommands commands, string sessionId, string? chartName, ChartAxisType? axis, string? title)
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
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' {axis!.Value} axis title set" }, JsonOptions);
    }

    private static string GetAxisNumberFormatAction(ChartCommands commands, string sessionId, string? chartName, ChartAxisType? axis)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "get-axis-number-format");
        if (!axis.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(axis), "get-axis-number-format");

        var numberFormat = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GetAxisNumberFormat(batch, chartName!, axis!.Value));

        return JsonSerializer.Serialize(new { Success = true, ChartName = chartName, Axis = axis!.Value.ToString(), NumberFormat = numberFormat }, JsonOptions);
    }

    private static string SetAxisNumberFormatAction(ChartCommands commands, string sessionId, string? chartName, ChartAxisType? axis, string? numberFormat)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-axis-number-format");
        if (!axis.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(axis), "set-axis-number-format");
        if (string.IsNullOrWhiteSpace(numberFormat))
            ExcelToolsBase.ThrowMissingParameter(nameof(numberFormat), "set-axis-number-format");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetAxisNumberFormat(batch, chartName!, axis!.Value, numberFormat!);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' {axis!.Value} axis number format set to '{numberFormat}'" }, JsonOptions);
    }

    private static string ShowLegendAction(ChartCommands commands, string sessionId, string? chartName, bool? visible, LegendPosition? legendPosition)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "show-legend");
        if (!visible.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(visible), "show-legend");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.ShowLegend(batch, chartName!, visible!.Value, legendPosition);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = visible!.Value ? $"Chart '{chartName}' legend shown" : $"Chart '{chartName}' legend hidden" }, JsonOptions);
    }

    private static string SetStyleAction(ChartCommands commands, string sessionId, string? chartName, int? styleId)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-style");
        if (!styleId.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(styleId), "set-style");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetStyle(batch, chartName!, styleId!.Value);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Chart '{chartName}' style set to {styleId!.Value}" }, JsonOptions);
    }
}
