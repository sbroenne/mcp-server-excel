using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
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

        var action = settings.Action.Trim().ToLowerInvariant();
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
        // Parse values from JSON string
        List<List<object?>>? values = null;
        if (!string.IsNullOrEmpty(settings.Values))
        {
            try
            {
                values = JsonSerializer.Deserialize<List<List<object?>>>(settings.Values, DaemonProtocol.JsonOptions);
            }
            catch
            {
                // If not valid JSON array, treat as single value
                values = [[settings.Values]];
            }
        }

        return new
        {
            sheetName = settings.SheetName,
            range = settings.Range,
            values
        };
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <NAME>")]
        public string? SheetName { get; init; }

        [CommandOption("--range <ADDRESS>")]
        public string? Range { get; init; }

        [CommandOption("--values <JSON>")]
        public string? Values { get; init; }
    }
}
