using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Sheet commands - thin wrapper that sends requests to daemon.
/// </summary>
internal sealed class SheetCommand : AsyncCommand<SheetCommand.Settings>
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

        var validActions = ActionValidator.GetValidActions<WorksheetAction>()
            .Concat(ActionValidator.GetValidActions<WorksheetStyleAction>())
            .ToArray();

        if (!ActionValidator.TryNormalizeAction(settings.Action, validActions, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"sheet.{action}";

        // Build args based on action
        object? args = action switch
        {
            "list" => null,
            "create" => new { sheetName = settings.SheetName },
            "rename" => new { sheetName = settings.SheetName, newName = settings.NewName },
            "delete" => new { sheetName = settings.SheetName },
            "copy" => new { sourceSheet = settings.SourceSheet, targetSheet = settings.TargetSheet },
            "move" => new { sheetName = settings.SheetName, beforeSheet = settings.BeforeSheet, afterSheet = settings.AfterSheet },
            _ => null
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

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (e.g., list, create, rename, delete, copy, move)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <NAME>")]
        [Description("Worksheet name")]
        public string? SheetName { get; init; }

        [CommandOption("--new-name <NAME>")]
        [Description("New name for rename operation")]
        public string? NewName { get; init; }

        [CommandOption("--source-sheet <NAME>")]
        [Description("Source worksheet for copy operation")]
        public string? SourceSheet { get; init; }

        [CommandOption("--target-sheet <NAME>")]
        [Description("Target worksheet name for copy operation")]
        public string? TargetSheet { get; init; }

        [CommandOption("--before-sheet <NAME>")]
        [Description("Place sheet before this sheet (move)")]
        public string? BeforeSheet { get; init; }

        [CommandOption("--after-sheet <NAME>")]
        [Description("Place sheet after this sheet (move)")]
        public string? AfterSheet { get; init; }
    }
}
