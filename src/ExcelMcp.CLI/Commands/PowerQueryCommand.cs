using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Service;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// PowerQuery commands - thin wrapper that sends requests to service.
/// Actions: list, view, create, update, rename, delete, refresh, refresh-all, load-to, get-load-config
/// </summary>
internal sealed class PowerQueryCommand : AsyncCommand<PowerQueryCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<PowerQueryAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"powerquery.{action}";

        // Note: property names must match daemon's Args classes (e.g., PowerQueryRenameArgs)
        var mCode = ResolveFileOrValue(settings.MCode, settings.MCodeFile);
        object? args = action switch
        {
            "list" => null,
            "view" => new { queryName = settings.QueryName },
            "create" => new { queryName = settings.QueryName, mCode, loadDestination = settings.LoadDestination, targetSheet = settings.TargetSheet, targetCellAddress = settings.TargetCell },
            "update" => new { queryName = settings.QueryName, mCode },
            "rename" => new { oldName = settings.QueryName, newName = settings.NewName },
            "delete" => new { queryName = settings.QueryName },
            "refresh" => new { queryName = settings.QueryName },
            "refresh-all" => null,
            "load-to" => new { queryName = settings.QueryName, loadDestination = settings.LoadDestination, targetSheet = settings.TargetSheet, targetCellAddress = settings.TargetCell },
            "get-load-config" => new { queryName = settings.QueryName },
            "unload" => new { queryName = settings.QueryName },
            "evaluate" => new { mCode },
            _ => new { queryName = settings.QueryName }
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
        [Description("The action to perform (e.g., list, view, create, update, refresh, evaluate)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--query <NAME>")]
        [Description("Power Query name")]
        public string? QueryName { get; init; }

        [CommandOption("--new-name <NAME>")]
        [Description("New name for rename operation")]
        public string? NewName { get; init; }

        [CommandOption("--mcode <CODE>")]
        [Description("Power Query M code formula")]
        public string? MCode { get; init; }

        [CommandOption("--mcode-file <PATH>")]
        [Description("Read M code from file instead of command line")]
        public string? MCodeFile { get; init; }

        [CommandOption("--load-destination <DEST>")]
        [Description("Load destination: worksheet, data-model, both, connection-only")]
        public string? LoadDestination { get; init; }

        [CommandOption("--target-sheet <NAME>")]
        [Description("Target worksheet for data load")]
        public string? TargetSheet { get; init; }

        [CommandOption("--target-cell <ADDRESS>")]
        [Description("Target cell address for data load (e.g., A1)")]
        public string? TargetCell { get; init; }
    }
}
