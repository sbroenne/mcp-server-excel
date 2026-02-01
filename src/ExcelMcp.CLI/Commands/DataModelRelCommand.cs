using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// DataModelRel commands - thin wrapper that sends relationship requests to daemon.
/// Actions: list-relationships, read-relationship, create-relationship, update-relationship, delete-relationship
/// </summary>
internal sealed class DataModelRelCommand : AsyncCommand<DataModelRelCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<DataModelRelAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"datamodelrel.{action}";

        // Note: property names must match daemon's Args classes
        object? args = action switch
        {
            "list-relationships" => null,
            "read-relationship" => new { fromTable = settings.FromTable, fromColumn = settings.FromColumn, toTable = settings.ToTable, toColumn = settings.ToColumn },
            "create-relationship" => new { fromTable = settings.FromTable, fromColumn = settings.FromColumn, toTable = settings.ToTable, toColumn = settings.ToColumn, active = settings.Active ?? true },
            "update-relationship" => new { fromTable = settings.FromTable, fromColumn = settings.FromColumn, toTable = settings.ToTable, toColumn = settings.ToColumn, active = settings.Active ?? true },
            "delete-relationship" => new { fromTable = settings.FromTable, fromColumn = settings.FromColumn, toTable = settings.ToTable, toColumn = settings.ToColumn },
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
        [Description("The action to perform (e.g., list-relationships, create-relationship)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--from-table <NAME>")]
        [Description("Source (many-side) table name containing the foreign key")]
        public string? FromTable { get; init; }

        [CommandOption("--from-column <NAME>")]
        [Description("Column in from-table that links to to-table")]
        public string? FromColumn { get; init; }

        [CommandOption("--to-table <NAME>")]
        [Description("Target (one-side/lookup) table name containing the primary key")]
        public string? ToTable { get; init; }

        [CommandOption("--to-column <NAME>")]
        [Description("Column in to-table that from-column links to (usually primary key)")]
        public string? ToColumn { get; init; }

        [CommandOption("--active")]
        [Description("Set relationship as active (default: true). Use --active false for inactive.")]
        [DefaultValue(true)]
        public bool? Active { get; init; }
    }
}
