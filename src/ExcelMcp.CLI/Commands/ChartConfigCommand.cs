using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Chart configuration commands - thin wrapper that sends requests to daemon.
/// Actions: set-source-range, add-series, remove-series, set-chart-type, set-title,
/// set-axis-title, get-axis-number-format, set-axis-number-format, show-legend,
/// set-style, set-data-labels, get-axis-scale, set-axis-scale, get-gridlines,
/// set-gridlines, set-series-format, list-trendlines, add-trendline, delete-trendline,
/// set-trendline, set-placement
/// </summary>
internal sealed class ChartConfigCommand : AsyncCommand<ChartConfigCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required. Use --session <id>[/]");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(settings.Action))
        {
            AnsiConsole.MarkupLine("[red]Action is required.[/]");
            return 1;
        }

        var action = settings.Action.Trim().ToLowerInvariant();
        var command = $"chartconfig.{action}";

        object? args = action switch
        {
            "set-source-range" => new { chartName = settings.ChartName, sourceRange = settings.SourceRange },
            "add-series" => new { chartName = settings.ChartName, seriesName = settings.SeriesName, valuesRange = settings.ValuesRange, categoryRange = settings.CategoryRange },
            "remove-series" => new { chartName = settings.ChartName, seriesIndex = settings.SeriesIndex },
            "set-chart-type" => new { chartName = settings.ChartName, chartType = settings.ChartType },
            "set-title" => new { chartName = settings.ChartName, title = settings.Title },
            "set-axis-title" => new { chartName = settings.ChartName, axis = settings.Axis, title = settings.Title },
            "get-axis-number-format" => new { chartName = settings.ChartName, axis = settings.Axis },
            "set-axis-number-format" => new { chartName = settings.ChartName, axis = settings.Axis, numberFormat = settings.NumberFormat },
            "show-legend" => new { chartName = settings.ChartName, visible = settings.Visible, legendPosition = settings.LegendPosition },
            "set-style" => new { chartName = settings.ChartName, styleId = settings.StyleId },
            "set-data-labels" => new
            {
                chartName = settings.ChartName,
                showValue = settings.ShowValue,
                showPercentage = settings.ShowPercentage,
                showSeriesName = settings.ShowSeriesName,
                showCategoryName = settings.ShowCategoryName,
                separator = settings.Separator,
                labelPosition = settings.LabelPosition,
                seriesIndex = settings.SeriesIndex
            },
            "get-axis-scale" => new { chartName = settings.ChartName, axis = settings.Axis },
            "set-axis-scale" => new
            {
                chartName = settings.ChartName,
                axis = settings.Axis,
                minimumScale = settings.MinimumScale,
                maximumScale = settings.MaximumScale,
                majorUnit = settings.MajorUnit,
                minorUnit = settings.MinorUnit
            },
            "get-gridlines" => new { chartName = settings.ChartName },
            "set-gridlines" => new { chartName = settings.ChartName, axis = settings.Axis, showMajor = settings.ShowMajor, showMinor = settings.ShowMinor },
            "set-series-format" => new
            {
                chartName = settings.ChartName,
                seriesIndex = settings.SeriesIndex,
                markerStyle = settings.MarkerStyle,
                markerSize = settings.MarkerSize,
                markerBackgroundColor = settings.MarkerBackgroundColor,
                markerForegroundColor = settings.MarkerForegroundColor
            },
            "list-trendlines" => new { chartName = settings.ChartName, seriesIndex = settings.SeriesIndex },
            "add-trendline" => new
            {
                chartName = settings.ChartName,
                seriesIndex = settings.SeriesIndex,
                trendlineType = settings.TrendlineType,
                displayEquation = settings.DisplayEquation,
                displayRSquared = settings.DisplayRSquared,
                trendlineName = settings.TrendlineName
            },
            "delete-trendline" => new { chartName = settings.ChartName, seriesIndex = settings.SeriesIndex, trendlineIndex = settings.TrendlineIndex },
            "set-trendline" => new
            {
                chartName = settings.ChartName,
                seriesIndex = settings.SeriesIndex,
                trendlineIndex = settings.TrendlineIndex,
                displayEquation = settings.DisplayEquation,
                displayRSquared = settings.DisplayRSquared,
                trendlineName = settings.TrendlineName
            },
            "set-placement" => new { chartName = settings.ChartName, placement = settings.Placement },
            _ => new { chartName = settings.ChartName }
        };

        using var client = new DaemonClient();
        var response = await client.SendAsync(new DaemonRequest
        {
            Command = command,
            SessionId = settings.SessionId,
            Args = args != null ? JsonSerializer.Serialize(args, DaemonProtocol.JsonOptions) : null
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(!string.IsNullOrEmpty(response.Result) ? response.Result : JsonSerializer.Serialize(new { success = true }, DaemonProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, DaemonProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--chart <NAME>")]
        public string? ChartName { get; init; }

        [CommandOption("--source-range <ADDRESS>")]
        public string? SourceRange { get; init; }

        [CommandOption("--series-name <NAME>")]
        public string? SeriesName { get; init; }

        [CommandOption("--values-range <ADDRESS>")]
        public string? ValuesRange { get; init; }

        [CommandOption("--category-range <ADDRESS>")]
        public string? CategoryRange { get; init; }

        [CommandOption("--series-index <INDEX>")]
        public int? SeriesIndex { get; init; }

        [CommandOption("--chart-type <TYPE>")]
        public string? ChartType { get; init; }

        [CommandOption("--title <TEXT>")]
        public string? Title { get; init; }

        [CommandOption("--axis <AXIS>")]
        public string? Axis { get; init; }

        [CommandOption("--number-format <FORMAT>")]
        public string? NumberFormat { get; init; }

        [CommandOption("--visible")]
        public bool? Visible { get; init; }

        [CommandOption("--legend-position <POSITION>")]
        public string? LegendPosition { get; init; }

        [CommandOption("--style-id <ID>")]
        public int? StyleId { get; init; }

        [CommandOption("--show-value")]
        public bool? ShowValue { get; init; }

        [CommandOption("--show-percentage")]
        public bool? ShowPercentage { get; init; }

        [CommandOption("--show-series-name")]
        public bool? ShowSeriesName { get; init; }

        [CommandOption("--show-category-name")]
        public bool? ShowCategoryName { get; init; }

        [CommandOption("--separator <TEXT>")]
        public string? Separator { get; init; }

        [CommandOption("--label-position <POSITION>")]
        public string? LabelPosition { get; init; }

        [CommandOption("--minimum-scale <VALUE>")]
        public double? MinimumScale { get; init; }

        [CommandOption("--maximum-scale <VALUE>")]
        public double? MaximumScale { get; init; }

        [CommandOption("--major-unit <VALUE>")]
        public double? MajorUnit { get; init; }

        [CommandOption("--minor-unit <VALUE>")]
        public double? MinorUnit { get; init; }

        [CommandOption("--show-major")]
        public bool? ShowMajor { get; init; }

        [CommandOption("--show-minor")]
        public bool? ShowMinor { get; init; }

        [CommandOption("--marker-style <STYLE>")]
        public string? MarkerStyle { get; init; }

        [CommandOption("--marker-size <SIZE>")]
        public int? MarkerSize { get; init; }

        [CommandOption("--marker-background-color <COLOR>")]
        public string? MarkerBackgroundColor { get; init; }

        [CommandOption("--marker-foreground-color <COLOR>")]
        public string? MarkerForegroundColor { get; init; }

        [CommandOption("--trendline-type <TYPE>")]
        public string? TrendlineType { get; init; }

        [CommandOption("--trendline-index <INDEX>")]
        public int? TrendlineIndex { get; init; }

        [CommandOption("--display-equation")]
        public bool? DisplayEquation { get; init; }

        [CommandOption("--display-rsquared")]
        public bool? DisplayRSquared { get; init; }

        [CommandOption("--trendline-name <NAME>")]
        public string? TrendlineName { get; init; }

        [CommandOption("--placement <MODE>")]
        public string? Placement { get; init; }
    }
}
