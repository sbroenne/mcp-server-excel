using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// VBA commands - thin wrapper that sends requests to daemon.
/// Actions: list, view, import, delete, run, update
/// </summary>
internal sealed class VbaCommand : AsyncCommand<VbaCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<VbaAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"vba.{action}";

        // Note: property names must match daemon's Args classes (e.g., VbaImportArgs, VbaRunArgs)
        object? args = action switch
        {
            "list" => null,
            "view" => new { moduleName = settings.ModuleName },
            "import" => new { moduleName = settings.ModuleName, vbaCode = settings.Code },
            "delete" => new { moduleName = settings.ModuleName },
            "run" => new { procedureName = settings.MacroName, parameters = ParseParameters(settings.Arguments) },
            "update" => new { moduleName = settings.ModuleName, vbaCode = settings.Code },
            _ => new { moduleName = settings.ModuleName }
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

    /// <summary>
    /// Parse comma-separated arguments string into a list.
    /// </summary>
    private static List<string>? ParseParameters(string? arguments)
    {
        if (string.IsNullOrWhiteSpace(arguments))
            return null;

        // Try to parse as JSON array first
        try
        {
            return JsonSerializer.Deserialize<List<string>>(arguments, DaemonProtocol.JsonOptions);
        }
        catch
        {
            // Fall back to comma-separated parsing
            return arguments.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--module <NAME>")]
        public string? ModuleName { get; init; }

        [CommandOption("--macro <NAME>")]
        public string? MacroName { get; init; }

        [CommandOption("--code <CODE>")]
        public string? Code { get; init; }

        [CommandOption("--module-type <TYPE>")]
        public string? ModuleType { get; init; }

        [CommandOption("--arguments <ARGS>")]
        public string? Arguments { get; init; }
    }
}
