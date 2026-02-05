using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Sbroenne.ExcelMcp.Generated;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Chart configuration commands - thin wrapper that sends requests to service.
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

        // Validate and normalize action
        var action = settings.Action.Trim().ToLowerInvariant();
        if (!ServiceRegistry.ChartConfig.ValidActions.Contains(action, StringComparer.OrdinalIgnoreCase))
        {
            var validList = string.Join(", ", ServiceRegistry.ChartConfig.ValidActions);
            AnsiConsole.MarkupLine($"[red]Invalid action '{action}'. Valid actions: {validList}[/]");
            return 1;
        }
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

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = settings.SessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null
        }, cancellationToken);

        if (response.Success)
        {
            Console.WriteLine(!string.IsNullOrEmpty(response.Result) ? response.Result : JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (e.g., set-title, add-series, set-data-labels)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--chart <NAME>")]
        [Description("Chart name")]
        public string? ChartName { get; init; }

        [CommandOption("--source-range <ADDRESS>")]
        [Description("Source data range address")]
        public string? SourceRange { get; init; }

        [CommandOption("--series-name <NAME>")]
        [Description("Name for new series")]
        public string? SeriesName { get; init; }

        [CommandOption("--values-range <ADDRESS>")]
        [Description("Range containing series values")]
        public string? ValuesRange { get; init; }

        [CommandOption("--category-range <ADDRESS>")]
        [Description("Range containing category labels")]
        public string? CategoryRange { get; init; }

        [CommandOption("--series-index <INDEX>")]
        [Description("1-based index of series to modify")]
        public int? SeriesIndex { get; init; }

        [CommandOption("--chart-type <TYPE>")]
        [Description("Chart type (e.g., ColumnClustered, Line, Pie)")]
        public string? ChartType { get; init; }

        [CommandOption("--title <TEXT>")]
        [Description("Chart or axis title text")]
        public string? Title { get; init; }

        [CommandOption("--axis <AXIS>")]
        [Description("Axis to configure (Category, Value, Series)")]
        public string? Axis { get; init; }

        [CommandOption("--number-format <FORMAT>")]
        [Description("Number format code (e.g., #,##0.00)")]
        public string? NumberFormat { get; init; }

        [CommandOption("--visible")]
        [Description("Show/hide element")]
        public bool? Visible { get; init; }

        [CommandOption("--legend-position <POSITION>")]
        [Description("Legend position (Top, Bottom, Left, Right, Corner)")]
        public string? LegendPosition { get; init; }

        [CommandOption("--style-id <ID>")]
        [Description("Built-in chart style ID (1-48)")]
        public int? StyleId { get; init; }

        [CommandOption("--show-value")]
        [Description("Show values in data labels")]
        public bool? ShowValue { get; init; }

        [CommandOption("--show-percentage")]
        [Description("Show percentages in data labels")]
        public bool? ShowPercentage { get; init; }

        [CommandOption("--show-series-name")]
        [Description("Show series name in data labels")]
        public bool? ShowSeriesName { get; init; }

        [CommandOption("--show-category-name")]
        [Description("Show category name in data labels")]
        public bool? ShowCategoryName { get; init; }

        [CommandOption("--separator <TEXT>")]
        [Description("Separator between data label parts")]
        public string? Separator { get; init; }

        [CommandOption("--label-position <POSITION>")]
        [Description("Data label position")]
        public string? LabelPosition { get; init; }

        [CommandOption("--minimum-scale <VALUE>")]
        [Description("Minimum axis scale value")]
        public double? MinimumScale { get; init; }

        [CommandOption("--maximum-scale <VALUE>")]
        [Description("Maximum axis scale value")]
        public double? MaximumScale { get; init; }

        [CommandOption("--major-unit <VALUE>")]
        [Description("Major gridline interval")]
        public double? MajorUnit { get; init; }

        [CommandOption("--minor-unit <VALUE>")]
        [Description("Minor gridline interval")]
        public double? MinorUnit { get; init; }

        [CommandOption("--show-major")]
        [Description("Show major gridlines")]
        public bool? ShowMajor { get; init; }

        [CommandOption("--show-minor")]
        [Description("Show minor gridlines")]
        public bool? ShowMinor { get; init; }

        [CommandOption("--marker-style <STYLE>")]
        [Description("Data point marker style")]
        public string? MarkerStyle { get; init; }

        [CommandOption("--marker-size <SIZE>")]
        [Description("Marker size in points")]
        public int? MarkerSize { get; init; }

        [CommandOption("--marker-background-color <COLOR>")]
        [Description("Marker fill color")]
        public string? MarkerBackgroundColor { get; init; }

        [CommandOption("--marker-foreground-color <COLOR>")]
        [Description("Marker border color")]
        public string? MarkerForegroundColor { get; init; }

        [CommandOption("--trendline-type <TYPE>")]
        [Description("Trendline type (Linear, Exponential, Power, etc.)")]
        public string? TrendlineType { get; init; }

        [CommandOption("--trendline-index <INDEX>")]
        [Description("1-based index of trendline")]
        public int? TrendlineIndex { get; init; }

        [CommandOption("--display-equation")]
        [Description("Display trendline equation on chart")]
        public bool? DisplayEquation { get; init; }

        [CommandOption("--display-rsquared")]
        [Description("Display R-squared value on chart")]
        public bool? DisplayRSquared { get; init; }

        [CommandOption("--trendline-name <NAME>")]
        [Description("Custom name for trendline")]
        public string? TrendlineName { get; init; }

        [CommandOption("--placement <MODE>")]
        [Description("Chart placement mode (FreeFloating, Move, MoveAndSize)")]
        public string? Placement { get; init; }
    }
}


