using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Table commands - thin wrapper that sends requests to daemon.
/// Actions: list, create, read, rename, delete, resize, set-style, toggle-totals,
/// set-column-total, append, get-data, add-to-datamodel, create-from-dax, update-dax, get-dax
/// </summary>
internal sealed class TableCommand : AsyncCommand<TableCommand.Settings>
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
        var command = $"table.{action}";

        // Build args based on action
        // Note: property names must match daemon's Args classes (e.g., TableCreateArgs)
        object? args = action switch
        {
            "list" => null,
            "create" => new { sheetName = settings.SheetName, tableName = settings.TableName, range = settings.Range, hasHeaders = settings.HasHeaders, tableStyle = settings.Style },
            "read" => new { tableName = settings.TableName },
            "rename" => new { tableName = settings.TableName, newName = settings.NewName },
            "delete" => new { tableName = settings.TableName },
            "resize" => new { tableName = settings.TableName, newRange = settings.Range },
            "set-style" => new { tableName = settings.TableName, tableStyle = settings.Style },
            "toggle-totals" => new { tableName = settings.TableName, showTotals = settings.HasHeaders },
            "set-column-total" => new { tableName = settings.TableName, columnName = settings.NewName, totalFunction = settings.Style },
            "append" => new { tableName = settings.TableName, csvData = settings.CsvData },
            "get-data" => new { tableName = settings.TableName, visibleOnly = settings.VisibleOnly },
            "add-to-datamodel" => new { tableName = settings.TableName },
            "create-from-dax" => new { sheetName = settings.SheetName, tableName = settings.TableName, daxQuery = settings.DaxQuery, targetCell = settings.Range },
            "update-dax" => new { tableName = settings.TableName, daxQuery = settings.DaxQuery },
            "get-dax" => new { tableName = settings.TableName },
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
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--table <NAME>")]
        public string? TableName { get; init; }

        [CommandOption("--sheet <NAME>")]
        public string? SheetName { get; init; }

        [CommandOption("--range <ADDRESS>")]
        public string? Range { get; init; }

        [CommandOption("--new-name <NAME>")]
        public string? NewName { get; init; }

        [CommandOption("--style <NAME>")]
        public string? Style { get; init; }

        [CommandOption("--has-headers")]
        public bool HasHeaders { get; init; } = true;

        [CommandOption("--csv-data <DATA>")]
        public string? CsvData { get; init; }

        [CommandOption("--visible-only")]
        public bool VisibleOnly { get; init; }

        [CommandOption("--dax-query <QUERY>")]
        public string? DaxQuery { get; init; }
    }
}
