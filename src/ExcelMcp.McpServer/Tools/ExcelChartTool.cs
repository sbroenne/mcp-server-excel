using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Chart operations
/// </summary>
[McpServerToolType]
public static class ExcelChartTool
{
    private static readonly JsonSerializerOptions JsonOptions = ExcelToolsBase.JsonOptions;

    [McpServerTool(Name = "excel_chart")]
    [Description(@"Excel Chart operations - create, modify, and analyze Regular Charts and PivotCharts.

⚡ CHART TYPES: Supports 70+ chart types including Column, Line, Pie, Bar, Area, XYScatter, and specialized types.

⚡ STRATEGY PATTERN: Automatically handles differences between Regular Charts (use SeriesCollection) and PivotCharts (read-only data source).")]
    public static string ExcelChart(
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        ChartAction action,

        [Description("Path to Excel file (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

        [Description("Chart name (required for most operations)")]
        string? chartName = null,

        [Description("Sheet name (for create-from-range)")]
        string? sheetName = null,

        [Description("Source data range (e.g., 'A1:D10' for create-from-range, set-source-range)")]
        string? sourceRange = null,

        [Description("Chart type enum value (e.g., ColumnClustered, Line, Pie, BarClustered, Area, XYScatter)")]
        ChartType? chartType = null,

        [Description("PivotTable name (for create-from-pivottable)")]
        string? pivotTableName = null,

        [Description("Left position in points")]
        double? left = null,

        [Description("Top position in points")]
        double? top = null,

        [Description("Width in points (optional, uses Excel default if not specified)")]
        double? width = null,

        [Description("Height in points (optional, uses Excel default if not specified)")]
        double? height = null,

        [Description("Chart title text (empty string hides title)")]
        string? title = null,

        [Description("Axis type for set-axis-title: Category, Value, CategorySecondary, ValueSecondary")]
        ChartAxisType? axis = null,

        [Description("Series name (for add-series)")]
        string? seriesName = null,

        [Description("Values range for series (e.g., 'Sheet1!B2:B10' for add-series)")]
        string? valuesRange = null,

        [Description("Category range for series (e.g., 'Sheet1!A2:A10', optional for add-series)")]
        string? categoryRange = null,

        [Description("1-based series index (for remove-series)")]
        int? seriesIndex = null,

        [Description("Show/hide legend: true=show, false=hide")]
        bool? visible = null,

        [Description("Legend position: Bottom, Corner, Top, Right, Left")]
        LegendPosition? legendPosition = null,

        [Description("Chart style ID (1-48)")]
        int? styleId = null)
    {
        return ExcelToolsBase.ExecuteToolAction(
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
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Charts,
            result.ErrorMessage
        }, JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Name,
            result.SheetName,
            result.ChartType,
            result.IsPivotChart,
            result.LinkedPivotTable,
            result.SourceRange,
            result.Left,
            result.Top,
            result.Width,
            result.Height,
            result.Title,
            result.HasLegend,
            result.Series,
            result.ErrorMessage
        }, JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ChartName,
            result.SheetName,
            result.ChartType,
            result.IsPivotChart,
            result.Left,
            result.Top,
            result.Width,
            result.Height,
            result.ErrorMessage
        }, JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ChartName,
            result.SheetName,
            result.ChartType,
            result.IsPivotChart,
            result.LinkedPivotTable,
            result.Left,
            result.Top,
            result.Width,
            result.Height,
            result.ErrorMessage
        }, JsonOptions);
    }

    private static string Delete(
        ChartCommands commands,
        string sessionId,
        string? chartName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "delete");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Delete(batch, chartName!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.Move(batch, chartName!, left, top, width, height));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetSourceRange(batch, chartName!, sourceRange!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SeriesName,
            result.ValuesRange,
            result.CategoryRange,
            result.SeriesIndex,
            result.ErrorMessage
        }, JsonOptions);
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.RemoveSeries(batch, chartName!, seriesIndex!.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetChartType(batch, chartName!, chartType!.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetTitle(batch, chartName!, title!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetAxisTitle(batch, chartName!, axis!.Value, title!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ShowLegend(batch, chartName!, visible!.Value, legendPosition));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
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

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.SetStyle(batch, chartName!, styleId!.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, JsonOptions);
    }
}
