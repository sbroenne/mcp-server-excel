using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.CLI.Infrastructure.Session;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands.Chart;

internal sealed class ChartCommand : Command<ChartCommand.Settings>
{
    private readonly ISessionService _sessionService;
    private readonly IChartCommands _chartCommands;
    private readonly ICliConsole _console;

    public ChartCommand(ISessionService sessionService, IChartCommands chartCommands, ICliConsole console)
    {
        _sessionService = sessionService ?? throw new ArgumentNullException(nameof(sessionService));
        _chartCommands = chartCommands ?? throw new ArgumentNullException(nameof(chartCommands));
        _console = console ?? throw new ArgumentNullException(nameof(console));
    }

    public override int Execute(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            _console.WriteError("Session ID is required. Use 'session open' first.");
            return -1;
        }

        var action = settings.Action?.Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(action))
        {
            _console.WriteError("Action is required.");
            return -1;
        }

        var batch = _sessionService.GetBatch(settings.SessionId);

        return action switch
        {
            "list" => ExecuteList(batch),
            "read" => ExecuteRead(batch, settings),
            "create-from-range" => ExecuteCreateFromRange(batch, settings),
            "create-from-pivottable" => ExecuteCreateFromPivotTable(batch, settings),
            "delete" => ExecuteDelete(batch, settings),
            "move" => ExecuteMove(batch, settings),
            "set-source-range" => ExecuteSetSourceRange(batch, settings),
            "add-series" => ExecuteAddSeries(batch, settings),
            "remove-series" => ExecuteRemoveSeries(batch, settings),
            "set-chart-type" => ExecuteSetChartType(batch, settings),
            "set-title" => ExecuteSetTitle(batch, settings),
            "set-axis-title" => ExecuteSetAxisTitle(batch, settings),
            "get-axis-number-format" => ExecuteGetAxisNumberFormat(batch, settings),
            "set-axis-number-format" => ExecuteSetAxisNumberFormat(batch, settings),
            "show-legend" => ExecuteShowLegend(batch, settings),
            "set-style" => ExecuteSetStyle(batch, settings),
            "set-data-labels" => ExecuteSetDataLabels(batch, settings),
            "get-axis-scale" => ExecuteGetAxisScale(batch, settings),
            "set-axis-scale" => ExecuteSetAxisScale(batch, settings),
            "get-gridlines" => ExecuteGetGridlines(batch, settings),
            "set-gridlines" => ExecuteSetGridlines(batch, settings),
            "set-series-format" => ExecuteSetSeriesFormat(batch, settings),
            "list-trendlines" => ExecuteListTrendlines(batch, settings),
            "add-trendline" => ExecuteAddTrendline(batch, settings),
            "delete-trendline" => ExecuteDeleteTrendline(batch, settings),
            "set-trendline" => ExecuteSetTrendline(batch, settings),
            "set-placement" => ExecuteSetPlacement(batch, settings),
            "fit-to-range" => ExecuteFitToRange(batch, settings),
            _ => ReportUnknown(action)
        };
    }

    private int ExecuteList(IExcelBatch batch)
    {
        try
        {
            var charts = _chartCommands.List(batch);
            _console.WriteJson(charts);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to list charts: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteRead(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName))
        {
            _console.WriteError("--chart-name is required for read.");
            return -1;
        }

        try
        {
            var result = _chartCommands.Read(batch, settings.ChartName);
            _console.WriteJson(result);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to read chart '{settings.ChartName}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteCreateFromRange(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.SheetName) ||
            string.IsNullOrWhiteSpace(settings.SourceRange) ||
            !settings.ChartType.HasValue ||
            !settings.Left.HasValue ||
            !settings.Top.HasValue)
        {
            _console.WriteError("--sheet, --source-range, --chart-type, --left, and --top are required for create-from-range.");
            return -1;
        }

        try
        {
            var result = _chartCommands.CreateFromRange(
                batch,
                settings.SheetName,
                settings.SourceRange,
                settings.ChartType.Value,
                settings.Left.Value,
                settings.Top.Value,
                settings.Width ?? 400,
                settings.Height ?? 300,
                settings.ChartName);
            _console.WriteJson(result);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to create chart from range: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteCreateFromPivotTable(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.PivotTableName) ||
            string.IsNullOrWhiteSpace(settings.SheetName) ||
            !settings.ChartType.HasValue ||
            !settings.Left.HasValue ||
            !settings.Top.HasValue)
        {
            _console.WriteError("--pivot-name, --sheet, --chart-type, --left, and --top are required for create-from-pivottable.");
            return -1;
        }

        try
        {
            var result = _chartCommands.CreateFromPivotTable(
                batch,
                settings.PivotTableName,
                settings.SheetName,
                settings.ChartType.Value,
                settings.Left.Value,
                settings.Top.Value,
                settings.Width ?? 400,
                settings.Height ?? 300,
                settings.ChartName);
            _console.WriteJson(result);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to create chart from PivotTable: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteDelete(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName))
        {
            _console.WriteError("--chart-name is required for delete.");
            return -1;
        }

        try
        {
            _chartCommands.Delete(batch, settings.ChartName);
            _console.WriteInfo($"Chart '{settings.ChartName}' deleted successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to delete chart '{settings.ChartName}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteMove(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName))
        {
            _console.WriteError("--chart-name is required for move.");
            return -1;
        }

        try
        {
            _chartCommands.Move(
                batch,
                settings.ChartName,
                settings.Left,
                settings.Top,
                settings.Width,
                settings.Height);
            _console.WriteInfo($"Chart '{settings.ChartName}' moved successfully.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to move chart '{settings.ChartName}': {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetSourceRange(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) ||
            string.IsNullOrWhiteSpace(settings.SourceRange))
        {
            _console.WriteError("--chart-name and --source-range are required for set-source-range.");
            return -1;
        }

        try
        {
            _chartCommands.SetSourceRange(batch, settings.ChartName, settings.SourceRange);
            _console.WriteInfo($"Chart '{settings.ChartName}' source range updated.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set source range: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteAddSeries(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) ||
            string.IsNullOrWhiteSpace(settings.SeriesName) ||
            string.IsNullOrWhiteSpace(settings.ValuesRange))
        {
            _console.WriteError("--chart-name, --series-name, and --values-range are required for add-series.");
            return -1;
        }

        try
        {
            var result = _chartCommands.AddSeries(
                batch,
                settings.ChartName,
                settings.SeriesName,
                settings.ValuesRange,
                settings.CategoryRange);
            _console.WriteJson(result);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to add series: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteRemoveSeries(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) ||
            !settings.SeriesIndex.HasValue)
        {
            _console.WriteError("--chart-name and --series-index are required for remove-series.");
            return -1;
        }

        try
        {
            _chartCommands.RemoveSeries(batch, settings.ChartName, settings.SeriesIndex.Value);
            _console.WriteInfo($"Series removed from chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to remove series: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetChartType(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) ||
            !settings.ChartType.HasValue)
        {
            _console.WriteError("--chart-name and --chart-type are required for set-chart-type.");
            return -1;
        }

        try
        {
            _chartCommands.SetChartType(batch, settings.ChartName, settings.ChartType.Value);
            _console.WriteInfo($"Chart type updated for '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set chart type: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetTitle(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName))
        {
            _console.WriteError("--chart-name is required for set-title.");
            return -1;
        }

        try
        {
            // Title can be empty string to hide
            var title = settings.Title ?? string.Empty;
            _chartCommands.SetTitle(batch, settings.ChartName, title);
            _console.WriteInfo($"Title updated for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set title: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetAxisTitle(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) ||
            !settings.AxisType.HasValue)
        {
            _console.WriteError("--chart-name and --axis-type are required for set-axis-title.");
            return -1;
        }

        try
        {
            // Title can be empty string to hide
            var title = settings.Title ?? string.Empty;
            _chartCommands.SetAxisTitle(batch, settings.ChartName, settings.AxisType.Value, title);
            _console.WriteInfo($"Axis title updated for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set axis title: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteShowLegend(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) ||
            !settings.Visible.HasValue)
        {
            _console.WriteError("--chart-name and --visible are required for show-legend.");
            return -1;
        }

        try
        {
            _chartCommands.ShowLegend(
                batch,
                settings.ChartName,
                settings.Visible.Value,
                settings.LegendPosition);
            _console.WriteInfo($"Legend updated for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to update legend: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetStyle(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) ||
            !settings.StyleId.HasValue)
        {
            _console.WriteError("--chart-name and --style-id are required for set-style.");
            return -1;
        }

        try
        {
            _chartCommands.SetStyle(batch, settings.ChartName, settings.StyleId.Value);
            _console.WriteInfo($"Style applied to chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set style: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetAxisNumberFormat(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.AxisType.HasValue)
        {
            _console.WriteError("--chart-name and --axis-type are required for get-axis-number-format.");
            return -1;
        }

        try
        {
            var format = _chartCommands.GetAxisNumberFormat(batch, settings.ChartName, settings.AxisType.Value);
            _console.WriteJson(new { success = true, chartName = settings.ChartName, axis = settings.AxisType.Value.ToString(), numberFormat = format });
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to get axis number format: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetAxisNumberFormat(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.AxisType.HasValue || string.IsNullOrWhiteSpace(settings.NumberFormat))
        {
            _console.WriteError("--chart-name, --axis-type, and --number-format are required for set-axis-number-format.");
            return -1;
        }

        try
        {
            _chartCommands.SetAxisNumberFormat(batch, settings.ChartName, settings.AxisType.Value, settings.NumberFormat);
            _console.WriteInfo($"Axis number format set for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set axis number format: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetDataLabels(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName))
        {
            _console.WriteError("--chart-name is required for set-data-labels.");
            return -1;
        }

        try
        {
            _chartCommands.SetDataLabels(
                batch,
                settings.ChartName,
                settings.ShowValue,
                settings.ShowPercentage,
                settings.ShowSeriesName,
                settings.ShowCategoryName,
                settings.ShowBubbleSize,
                settings.Separator,
                settings.LabelPosition,
                settings.SeriesIndex);
            _console.WriteInfo($"Data labels updated for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set data labels: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetAxisScale(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.AxisType.HasValue)
        {
            _console.WriteError("--chart-name and --axis-type are required for get-axis-scale.");
            return -1;
        }

        try
        {
            var result = _chartCommands.GetAxisScale(batch, settings.ChartName, settings.AxisType.Value);
            _console.WriteJson(result);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to get axis scale: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetAxisScale(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.AxisType.HasValue)
        {
            _console.WriteError("--chart-name and --axis-type are required for set-axis-scale.");
            return -1;
        }

        try
        {
            _chartCommands.SetAxisScale(
                batch,
                settings.ChartName,
                settings.AxisType.Value,
                settings.MinimumScale,
                settings.MaximumScale,
                settings.MajorUnit,
                settings.MinorUnit);
            _console.WriteInfo($"Axis scale set for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set axis scale: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteGetGridlines(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName))
        {
            _console.WriteError("--chart-name is required for get-gridlines.");
            return -1;
        }

        try
        {
            var result = _chartCommands.GetGridlines(batch, settings.ChartName);
            _console.WriteJson(result);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to get gridlines: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetGridlines(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.AxisType.HasValue)
        {
            _console.WriteError("--chart-name and --axis-type are required for set-gridlines.");
            return -1;
        }

        try
        {
            _chartCommands.SetGridlines(
                batch,
                settings.ChartName,
                settings.AxisType.Value,
                settings.ShowMajor,
                settings.ShowMinor);
            _console.WriteInfo($"Gridlines updated for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set gridlines: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetSeriesFormat(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.SeriesIndex.HasValue)
        {
            _console.WriteError("--chart-name and --series-index are required for set-series-format.");
            return -1;
        }

        try
        {
            _chartCommands.SetSeriesFormat(
                batch,
                settings.ChartName,
                settings.SeriesIndex.Value,
                settings.MarkerStyle,
                settings.MarkerSize,
                settings.MarkerBackgroundColor,
                settings.MarkerForegroundColor,
                settings.InvertIfNegative);
            _console.WriteInfo($"Series format updated for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set series format: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteListTrendlines(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.SeriesIndex.HasValue)
        {
            _console.WriteError("--chart-name and --series-index are required for list-trendlines.");
            return -1;
        }

        try
        {
            var result = _chartCommands.ListTrendlines(batch, settings.ChartName, settings.SeriesIndex.Value);
            _console.WriteJson(result);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to list trendlines: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteAddTrendline(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.SeriesIndex.HasValue || !settings.TrendlineType.HasValue)
        {
            _console.WriteError("--chart-name, --series-index, and --trendline-type are required for add-trendline.");
            return -1;
        }

        try
        {
            var result = _chartCommands.AddTrendline(
                batch,
                settings.ChartName,
                settings.SeriesIndex.Value,
                settings.TrendlineType.Value,
                settings.TrendlineOrder,
                settings.TrendlinePeriod,
                settings.TrendlineForward,
                settings.TrendlineBackward,
                settings.TrendlineIntercept,
                settings.DisplayEquation ?? false,
                settings.DisplayRSquared ?? false,
                settings.TrendlineName);
            _console.WriteJson(result);
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to add trendline: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteDeleteTrendline(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.SeriesIndex.HasValue || !settings.TrendlineIndex.HasValue)
        {
            _console.WriteError("--chart-name, --series-index, and --trendline-index are required for delete-trendline.");
            return -1;
        }

        try
        {
            _chartCommands.DeleteTrendline(batch, settings.ChartName, settings.SeriesIndex.Value, settings.TrendlineIndex.Value);
            _console.WriteInfo($"Trendline deleted from chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to delete trendline: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetTrendline(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.SeriesIndex.HasValue || !settings.TrendlineIndex.HasValue)
        {
            _console.WriteError("--chart-name, --series-index, and --trendline-index are required for set-trendline.");
            return -1;
        }

        try
        {
            _chartCommands.SetTrendline(
                batch,
                settings.ChartName,
                settings.SeriesIndex.Value,
                settings.TrendlineIndex.Value,
                settings.TrendlineForward,
                settings.TrendlineBackward,
                settings.TrendlineIntercept,
                settings.DisplayEquation,
                settings.DisplayRSquared,
                settings.TrendlineName);
            _console.WriteInfo($"Trendline updated for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set trendline: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteSetPlacement(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || !settings.Placement.HasValue)
        {
            _console.WriteError("--chart-name and --placement are required for set-placement. (1=Move and size with cells, 2=Move but don't size, 3=Don't move or size)");
            return -1;
        }

        try
        {
            _chartCommands.SetPlacement(batch, settings.ChartName, settings.Placement.Value);
            _console.WriteInfo($"Placement mode set for chart '{settings.ChartName}'.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to set placement: {ex.Message}");
            return 1;
        }
    }

    private int ExecuteFitToRange(IExcelBatch batch, Settings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.ChartName) || string.IsNullOrWhiteSpace(settings.SheetName) || string.IsNullOrWhiteSpace(settings.SourceRange))
        {
            _console.WriteError("--chart-name, --sheet, and --source-range are required for fit-to-range.");
            return -1;
        }

        try
        {
            _chartCommands.FitToRange(batch, settings.ChartName, settings.SheetName, settings.SourceRange);
            _console.WriteInfo($"Chart '{settings.ChartName}' fitted to range.");
            return 0;
        }
        catch (Exception ex)
        {
            _console.WriteError($"Failed to fit chart to range: {ex.Message}");
            return 1;
        }
    }

    private int ReportUnknown(string action)
    {
        _console.WriteError($"Unknown Chart action '{action}'.");
        return -1;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<action>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--chart-name <NAME>")]
        public string? ChartName { get; init; }

        [CommandOption("--sheet <SHEET>")]
        public string? SheetName { get; init; }

        [CommandOption("--source-range <RANGE>")]
        public string? SourceRange { get; init; }

        [CommandOption("--chart-type <TYPE>")]
        public ChartType? ChartType { get; init; }

        [CommandOption("--pivot-name <NAME>")]
        public string? PivotTableName { get; init; }

        [CommandOption("--left <POINTS>")]
        public double? Left { get; init; }

        [CommandOption("--top <POINTS>")]
        public double? Top { get; init; }

        [CommandOption("--width <POINTS>")]
        public double? Width { get; init; }

        [CommandOption("--height <POINTS>")]
        public double? Height { get; init; }

        [CommandOption("--title <TEXT>")]
        public string? Title { get; init; }

        [CommandOption("--axis-type <TYPE>")]
        public ChartAxisType? AxisType { get; init; }

        [CommandOption("--series-name <NAME>")]
        public string? SeriesName { get; init; }

        [CommandOption("--values-range <RANGE>")]
        public string? ValuesRange { get; init; }

        [CommandOption("--category-range <RANGE>")]
        public string? CategoryRange { get; init; }

        [CommandOption("--series-index <INDEX>")]
        public int? SeriesIndex { get; init; }

        [CommandOption("--visible <BOOL>")]
        public bool? Visible { get; init; }

        [CommandOption("--legend-position <POSITION>")]
        public LegendPosition? LegendPosition { get; init; }

        [CommandOption("--style-id <ID>")]
        public int? StyleId { get; init; }

        // Data labels options
        [CommandOption("--show-value <BOOL>")]
        public bool? ShowValue { get; init; }

        [CommandOption("--show-percentage <BOOL>")]
        public bool? ShowPercentage { get; init; }

        [CommandOption("--show-series-name <BOOL>")]
        public bool? ShowSeriesName { get; init; }

        [CommandOption("--show-category-name <BOOL>")]
        public bool? ShowCategoryName { get; init; }

        [CommandOption("--show-bubble-size <BOOL>")]
        public bool? ShowBubbleSize { get; init; }

        [CommandOption("--separator <TEXT>")]
        public string? Separator { get; init; }

        [CommandOption("--label-position <POSITION>")]
        public DataLabelPosition? LabelPosition { get; init; }

        // Axis scale options
        [CommandOption("--minimum-scale <VALUE>")]
        public double? MinimumScale { get; init; }

        [CommandOption("--maximum-scale <VALUE>")]
        public double? MaximumScale { get; init; }

        [CommandOption("--major-unit <VALUE>")]
        public double? MajorUnit { get; init; }

        [CommandOption("--minor-unit <VALUE>")]
        public double? MinorUnit { get; init; }

        [CommandOption("--number-format <FORMAT>")]
        public string? NumberFormat { get; init; }

        // Gridlines options
        [CommandOption("--show-major <BOOL>")]
        public bool? ShowMajor { get; init; }

        [CommandOption("--show-minor <BOOL>")]
        public bool? ShowMinor { get; init; }

        // Series format options
        [CommandOption("--marker-style <STYLE>")]
        public MarkerStyle? MarkerStyle { get; init; }

        [CommandOption("--marker-size <SIZE>")]
        public int? MarkerSize { get; init; }

        [CommandOption("--marker-background-color <HEX>")]
        public string? MarkerBackgroundColor { get; init; }

        [CommandOption("--marker-foreground-color <HEX>")]
        public string? MarkerForegroundColor { get; init; }

        [CommandOption("--invert-if-negative <BOOL>")]
        public bool? InvertIfNegative { get; init; }

        // Trendline options
        [CommandOption("--trendline-type <TYPE>")]
        public TrendlineType? TrendlineType { get; init; }

        [CommandOption("--trendline-index <INDEX>")]
        public int? TrendlineIndex { get; init; }

        [CommandOption("--trendline-order <ORDER>")]
        public int? TrendlineOrder { get; init; }

        [CommandOption("--trendline-period <PERIOD>")]
        public int? TrendlinePeriod { get; init; }

        [CommandOption("--trendline-forward <PERIODS>")]
        public double? TrendlineForward { get; init; }

        [CommandOption("--trendline-backward <PERIODS>")]
        public double? TrendlineBackward { get; init; }

        [CommandOption("--trendline-intercept <VALUE>")]
        public double? TrendlineIntercept { get; init; }

        [CommandOption("--display-equation <BOOL>")]
        public bool? DisplayEquation { get; init; }

        [CommandOption("--display-r-squared <BOOL>")]
        public bool? DisplayRSquared { get; init; }

        [CommandOption("--trendline-name <NAME>")]
        public string? TrendlineName { get; init; }

        // Placement option
        [CommandOption("--placement <MODE>")]
        public int? Placement { get; init; }
    }
}
