using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// DataModel commands - thin wrapper that sends requests to daemon.
/// Actions: list-tables, read-table, list-columns, list-measures, read, create-measure, update-measure, delete-measure,
/// rename-table, delete-table, read-info, refresh, evaluate, execute-dmv
/// </summary>
internal sealed class DataModelCommand : AsyncCommand<DataModelCommand.Settings>
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

        if (!ActionValidator.TryNormalizeAction<DataModelAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"datamodel.{action}";

        // Note: property names must match daemon's Args classes (e.g., DataModelCreateMeasureArgs)
        object? args = action switch
        {
            "list-tables" => null,
            "read-table" => new { tableName = settings.TableName, maxRows = settings.MaxRows },
            "list-columns" => new { tableName = settings.TableName },
            "list-measures" => new { tableName = settings.TableName },
            "read" => new { measureName = settings.MeasureName, tableName = settings.TableName },
            "create-measure" => new { tableName = settings.TableName, measureName = settings.MeasureName, daxFormula = settings.Expression, formatType = settings.FormatString },
            "update-measure" => new { measureName = settings.MeasureName, tableName = settings.TableName, daxFormula = settings.Expression, formatType = settings.FormatString },
            "delete-measure" => new { measureName = settings.MeasureName, tableName = settings.TableName },
            "rename-table" => new { oldName = settings.TableName, newName = settings.NewName },
            "delete-table" => new { tableName = settings.TableName },
            "read-info" => null,
            "refresh" => null,
            "evaluate" => new { daxQuery = settings.DaxQuery },
            "execute-dmv" => new { dmvQuery = settings.DmvQuery },
            _ => new { tableName = settings.TableName }
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

        [CommandOption("--table <NAME>")]
        public string? TableName { get; init; }

        [CommandOption("--measure <NAME>")]
        public string? MeasureName { get; init; }

        [CommandOption("--new-name <NAME>")]
        public string? NewName { get; init; }

        [CommandOption("--expression <DAX>")]
        public string? Expression { get; init; }

        [CommandOption("--format-string <FORMAT>")]
        public string? FormatString { get; init; }

        [CommandOption("--dax-query <QUERY>")]
        public string? DaxQuery { get; init; }

        [CommandOption("--dmv-query <QUERY>")]
        public string? DmvQuery { get; init; }

        [CommandOption("--max-rows <COUNT>")]
        public int? MaxRows { get; init; }
    }
}
