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
    /// - get-axis-scale: Get min/max scale and unit settings for axis
    /// - set-axis-scale: Set min/max scale and major/minor unit values for axis
    ///
    /// LEGEND POSITIONS: Bottom, Corner, Top, Right, Left
    ///
    /// CHART STYLES: 1-48 (built-in Excel styles with different color schemes)
    ///
    /// DATA LABELS (set-data-labels):
    /// - Show values, percentages, series names, category names
    /// - Specify position (Center, InsideEnd, InsideBase, OutsideEnd, BestFit)
    /// - Apply to all series or specific series by index
    ///
    /// GRIDLINES (get-gridlines, set-gridlines):
    /// - Control major/minor gridlines for value and category axes
    ///
    /// SERIES FORMATTING (set-series-format):
    /// - Marker style (None, Circle, Square, Diamond, Triangle, X, Star, Plus)
    /// - Marker size (2-72 points)
    /// - Marker colors (#RRGGBB hex)
    /// - Invert if negative
    ///
    /// TRENDLINES:
    /// - list-trendlines: List all trendlines on a series
    /// - add-trendline: Add trendline (Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage)
    /// - delete-trendline: Remove a trendline by index
    /// - set-trendline: Configure trendline display options
    ///
    /// TRENDLINE TYPES:
    /// - Linear: Linear regression (y = mx + b)
    /// - Exponential: Exponential curve (y = ce^bx)
    /// - Logarithmic: Logarithmic curve (y = c ln x + b)
    /// - Polynomial: Polynomial fit (requires order 2-6)
    /// - Power: Power curve (y = cx^b)
    /// - MovingAverage: Moving average (requires period)
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
    /// <param name="seriesIndex">1-based series index for remove-series, set-data-labels, set-series-format actions</param>
    /// <param name="visible">Show or hide legend in show-legend action</param>
    /// <param name="legendPosition">Legend position: Bottom, Corner, Top, Right, Left</param>
    /// <param name="styleId">Chart style ID from 1-48 for set-style action</param>
    /// <param name="showValue">For set-data-labels: Show actual data values</param>
    /// <param name="showPercentage">For set-data-labels: Show percentage (pie/doughnut charts)</param>
    /// <param name="showSeriesName">For set-data-labels: Show series name in label</param>
    /// <param name="showCategoryName">For set-data-labels: Show category name in label</param>
    /// <param name="showBubbleSize">For set-data-labels: Show bubble size (bubble charts)</param>
    /// <param name="separator">For set-data-labels: Separator between label parts (e.g., ", " or "\n")</param>
    /// <param name="labelPosition">For set-data-labels: Position (Center, InsideEnd, InsideBase, OutsideEnd, BestFit, Above, Below, Left, Right)</param>
    /// <param name="minimumScale">For set-axis-scale: Minimum axis value (omit for auto)</param>
    /// <param name="maximumScale">For set-axis-scale: Maximum axis value (omit for auto)</param>
    /// <param name="majorUnit">For set-axis-scale: Major unit interval (omit for auto)</param>
    /// <param name="minorUnit">For set-axis-scale: Minor unit interval (omit for auto)</param>
    /// <param name="showMajor">For set-gridlines: Show major gridlines</param>
    /// <param name="showMinor">For set-gridlines: Show minor gridlines</param>
    /// <param name="markerStyle">For set-series-format: Marker style (None, Circle, Square, Diamond, Triangle, X, Star, Plus, etc.)</param>
    /// <param name="markerSize">For set-series-format: Marker size in points (2-72)</param>
    /// <param name="markerBackgroundColor">For set-series-format: Marker fill color (#RRGGBB hex)</param>
    /// <param name="markerForegroundColor">For set-series-format: Marker border color (#RRGGBB hex)</param>
    /// <param name="invertIfNegative">For set-series-format: Invert colors for negative values</param>
    /// <param name="trendlineType">For add-trendline: Type (Linear, Exponential, Logarithmic, Polynomial, Power, MovingAverage)</param>
    /// <param name="trendlineIndex">For delete-trendline, set-trendline: 1-based trendline index within the series</param>
    /// <param name="order">For add-trendline: Polynomial order (2-6) when type is Polynomial</param>
    /// <param name="period">For add-trendline: Moving average period when type is MovingAverage</param>
    /// <param name="forward">For add-trendline, set-trendline: Periods to forecast forward</param>
    /// <param name="backward">For add-trendline, set-trendline: Periods to forecast backward</param>
    /// <param name="intercept">For add-trendline, set-trendline: Y-intercept value (omit for calculated)</param>
    /// <param name="displayEquation">For add-trendline, set-trendline: Display equation on chart</param>
    /// <param name="displayRSquared">For add-trendline, set-trendline: Display R-squared value on chart</param>
    /// <param name="trendlineName">For add-trendline, set-trendline: Custom name for the trendline</param>
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
        [DefaultValue(null)] int? styleId,
        // Data labels parameters
        [DefaultValue(null)] bool? showValue,
        [DefaultValue(null)] bool? showPercentage,
        [DefaultValue(null)] bool? showSeriesName,
        [DefaultValue(null)] bool? showCategoryName,
        [DefaultValue(null)] bool? showBubbleSize,
        [DefaultValue(null)] string? separator,
        [DefaultValue(null)] DataLabelPosition? labelPosition,
        // Axis scale parameters
        [DefaultValue(null)] double? minimumScale,
        [DefaultValue(null)] double? maximumScale,
        [DefaultValue(null)] double? majorUnit,
        [DefaultValue(null)] double? minorUnit,
        // Gridlines parameters
        [DefaultValue(null)] bool? showMajor,
        [DefaultValue(null)] bool? showMinor,
        // Series format parameters
        [DefaultValue(null)] MarkerStyle? markerStyle,
        [DefaultValue(null)] int? markerSize,
        [DefaultValue(null)] string? markerBackgroundColor,
        [DefaultValue(null)] string? markerForegroundColor,
        [DefaultValue(null)] bool? invertIfNegative,
        // Trendline parameters
        [DefaultValue(null)] TrendlineType? trendlineType,
        [DefaultValue(null)] int? trendlineIndex,
        [DefaultValue(null)] int? order,
        [DefaultValue(null)] int? period,
        [DefaultValue(null)] double? forward,
        [DefaultValue(null)] double? backward,
        [DefaultValue(null)] double? intercept,
        [DefaultValue(null)] bool? displayEquation,
        [DefaultValue(null)] bool? displayRSquared,
        [DefaultValue(null)] string? trendlineName)
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
                    ChartConfigAction.SetDataLabels => SetDataLabelsAction(commands, sessionId, chartName, showValue, showPercentage, showSeriesName, showCategoryName, showBubbleSize, separator, labelPosition, seriesIndex),
                    ChartConfigAction.GetAxisScale => GetAxisScaleAction(commands, sessionId, chartName, axis),
                    ChartConfigAction.SetAxisScale => SetAxisScaleAction(commands, sessionId, chartName, axis, minimumScale, maximumScale, majorUnit, minorUnit),
                    ChartConfigAction.GetGridlines => GetGridlinesAction(commands, sessionId, chartName),
                    ChartConfigAction.SetGridlines => SetGridlinesAction(commands, sessionId, chartName, axis, showMajor, showMinor),
                    ChartConfigAction.SetSeriesFormat => SetSeriesFormatAction(commands, sessionId, chartName, seriesIndex, markerStyle, markerSize, markerBackgroundColor, markerForegroundColor, invertIfNegative),
                    ChartConfigAction.ListTrendlines => ListTrendlinesAction(commands, sessionId, chartName, seriesIndex),
                    ChartConfigAction.AddTrendline => AddTrendlineAction(commands, sessionId, chartName, seriesIndex, trendlineType, order, period, forward, backward, intercept, displayEquation, displayRSquared, trendlineName),
                    ChartConfigAction.DeleteTrendline => DeleteTrendlineAction(commands, sessionId, chartName, seriesIndex, trendlineIndex),
                    ChartConfigAction.SetTrendline => SetTrendlineAction(commands, sessionId, chartName, seriesIndex, trendlineIndex, forward, backward, intercept, displayEquation, displayRSquared, trendlineName),
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

    // === DATA LABELS ===

    private static string SetDataLabelsAction(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        bool? showValue,
        bool? showPercentage,
        bool? showSeriesName,
        bool? showCategoryName,
        bool? showBubbleSize,
        string? separator,
        DataLabelPosition? labelPosition,
        int? seriesIndex)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-data-labels");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetDataLabels(
                batch,
                chartName!,
                showValue,
                showPercentage,
                showSeriesName,
                showCategoryName,
                showBubbleSize,
                separator,
                labelPosition,
                seriesIndex);
            return 0;
        });

        string target = seriesIndex.HasValue ? $"series {seriesIndex.Value}" : "all series";
        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Data labels configured for {target} in chart '{chartName}'" }, JsonOptions);
    }

    // === AXIS SCALE ===

    private static string GetAxisScaleAction(ChartCommands commands, string sessionId, string? chartName, ChartAxisType? axis)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "get-axis-scale");
        if (!axis.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(axis), "get-axis-scale");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GetAxisScale(batch, chartName!, axis!.Value));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string SetAxisScaleAction(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        ChartAxisType? axis,
        double? minimumScale,
        double? maximumScale,
        double? majorUnit,
        double? minorUnit)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-axis-scale");
        if (!axis.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(axis), "set-axis-scale");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetAxisScale(batch, chartName!, axis!.Value, minimumScale, maximumScale, majorUnit, minorUnit);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Axis scale configured for {axis!.Value} axis in chart '{chartName}'" }, JsonOptions);
    }

    // === GRIDLINES ===

    private static string GetGridlinesAction(ChartCommands commands, string sessionId, string? chartName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "get-gridlines");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.GetGridlines(batch, chartName!));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string SetGridlinesAction(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        ChartAxisType? axis,
        bool? showMajor,
        bool? showMinor)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-gridlines");
        if (!axis.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(axis), "set-gridlines");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetGridlines(batch, chartName!, axis!.Value, showMajor, showMinor);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Gridlines configured for {axis!.Value} axis in chart '{chartName}'" }, JsonOptions);
    }

    // === SERIES FORMATTING ===

    private static string SetSeriesFormatAction(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        int? seriesIndex,
        MarkerStyle? markerStyle,
        int? markerSize,
        string? markerBackgroundColor,
        string? markerForegroundColor,
        bool? invertIfNegative)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-series-format");
        if (!seriesIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(seriesIndex), "set-series-format");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetSeriesFormat(
                batch,
                chartName!,
                seriesIndex!.Value,
                markerStyle,
                markerSize,
                markerBackgroundColor,
                markerForegroundColor,
                invertIfNegative);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Series {seriesIndex!.Value} format configured in chart '{chartName}'" }, JsonOptions);
    }

    // === TRENDLINES ===

    private static string ListTrendlinesAction(ChartCommands commands, string sessionId, string? chartName, int? seriesIndex)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "list-trendlines");
        if (!seriesIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(seriesIndex), "list-trendlines");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.ListTrendlines(batch, chartName!, seriesIndex!.Value));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string AddTrendlineAction(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        int? seriesIndex,
        TrendlineType? trendlineType,
        int? order,
        int? period,
        double? forward,
        double? backward,
        double? intercept,
        bool? displayEquation,
        bool? displayRSquared,
        string? trendlineName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "add-trendline");
        if (!seriesIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(seriesIndex), "add-trendline");
        if (!trendlineType.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(trendlineType), "add-trendline");

        var result = ExcelToolsBase.WithSession(sessionId,
            batch => commands.AddTrendline(
                batch,
                chartName!,
                seriesIndex!.Value,
                trendlineType!.Value,
                order,
                period,
                forward,
                backward,
                intercept,
                displayEquation ?? false,
                displayRSquared ?? false,
                trendlineName));

        return JsonSerializer.Serialize(result, JsonOptions);
    }

    private static string DeleteTrendlineAction(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        int? seriesIndex,
        int? trendlineIndex)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "delete-trendline");
        if (!seriesIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(seriesIndex), "delete-trendline");
        if (!trendlineIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(trendlineIndex), "delete-trendline");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.DeleteTrendline(batch, chartName!, seriesIndex!.Value, trendlineIndex!.Value);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Trendline {trendlineIndex!.Value} deleted from series {seriesIndex!.Value} in chart '{chartName}'" }, JsonOptions);
    }

    private static string SetTrendlineAction(
        ChartCommands commands,
        string sessionId,
        string? chartName,
        int? seriesIndex,
        int? trendlineIndex,
        double? forward,
        double? backward,
        double? intercept,
        bool? displayEquation,
        bool? displayRSquared,
        string? trendlineName)
    {
        if (string.IsNullOrWhiteSpace(chartName))
            ExcelToolsBase.ThrowMissingParameter(nameof(chartName), "set-trendline");
        if (!seriesIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(seriesIndex), "set-trendline");
        if (!trendlineIndex.HasValue)
            ExcelToolsBase.ThrowMissingParameter(nameof(trendlineIndex), "set-trendline");

        ExcelToolsBase.WithSession(sessionId, batch =>
        {
            commands.SetTrendline(
                batch,
                chartName!,
                seriesIndex!.Value,
                trendlineIndex!.Value,
                forward,
                backward,
                intercept,
                displayEquation,
                displayRSquared,
                trendlineName);
            return 0;
        });

        return JsonSerializer.Serialize(new OperationResult { Success = true, Message = $"Trendline {trendlineIndex!.Value} configured in series {seriesIndex!.Value} of chart '{chartName}'" }, JsonOptions);
    }
}
