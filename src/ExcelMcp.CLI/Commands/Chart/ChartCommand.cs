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
            "show-legend" => ExecuteShowLegend(batch, settings),
            "set-style" => ExecuteSetStyle(batch, settings),
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
    }
}
