using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Range commands - thin wrapper that sends requests to daemon.
/// </summary>
internal sealed class RangeCommand : AsyncCommand<RangeCommand.Settings>
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

        var validActions = ActionValidator.GetValidActions<RangeAction>()
            .Concat(ActionValidator.GetValidActions<RangeEditAction>())
            .Concat(ActionValidator.GetValidActions<RangeFormatAction>())
            .Concat(ActionValidator.GetValidActions<RangeLinkAction>())
            .ToArray();

        if (!ActionValidator.TryNormalizeAction(settings.Action, validActions, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"range.{action}";

        // Build args based on action
        object? args = action switch
        {
            "get-values" => new { sheetName = settings.SheetName, range = settings.Range },
            "set-values" => BuildSetValuesArgs(settings),
            "get-used-range" => new { sheetName = settings.SheetName },
            _ => new { sheetName = settings.SheetName, range = settings.Range }
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
            if (!string.IsNullOrEmpty(response.Result))
            {
                Console.WriteLine(response.Result);
            }
            else
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = true }, DaemonProtocol.JsonOptions));
            }
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, DaemonProtocol.JsonOptions));
            return 1;
        }
    }

    private static object? BuildSetValuesArgs(Settings settings)
    {
        // Parse values from JSON string or file
        List<List<object?>>? values = null;
        var valuesJson = ResolveFileOrValue(settings.Values, settings.ValuesFile);
        if (!string.IsNullOrEmpty(valuesJson))
        {
            try
            {
                values = JsonSerializer.Deserialize<List<List<object?>>>(valuesJson, DaemonProtocol.JsonOptions);
            }
            catch
            {
                // If not valid JSON array, treat as single value
                values = [[valuesJson]];
            }
        }

        return new
        {
            sheetName = settings.SheetName,
            range = settings.Range,
            values
        };
    }

    /// <summary>
    /// Returns file contents if filePath is provided, otherwise returns the direct value.
    /// </summary>
    private static string? ResolveFileOrValue(string? directValue, string? filePath)
    {
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found: {filePath}");
            }
            return File.ReadAllText(filePath);
        }
        return directValue;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (e.g., get-values, set-values, get-used-range)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <NAME>")]
        [Description("Target worksheet name")]
        public string? SheetName { get; init; }

        [CommandOption("--range <ADDRESS>")]
        [Description("Cell range address (e.g., A1:C10)")]
        public string? Range { get; init; }

        [CommandOption("--values <JSON>")]
        [Description("Cell values as JSON 2D array")]
        public string? Values { get; init; }

        [CommandOption("--values-file <PATH>")]
        [Description("Read values JSON from file instead of command line")]
        public string? ValuesFile { get; init; }
    }
}
