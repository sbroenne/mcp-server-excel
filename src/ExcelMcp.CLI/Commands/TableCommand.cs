using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Daemon;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
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

        if (!ActionValidator.TryNormalizeAction<TableAction>(settings.Action, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"table.{action}";

        // Build args based on action
        // Note: property names must match daemon's Args classes (e.g., TableCreateArgs)
        var csvData = ResolveFileOrValue(settings.CsvData, settings.CsvFile);
        var daxQuery = ResolveFileOrValue(settings.DaxQuery, settings.DaxQueryFile);
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
            "append" => new { tableName = settings.TableName, csvData },
            "get-data" => new { tableName = settings.TableName, visibleOnly = settings.VisibleOnly },
            "add-to-datamodel" => new { tableName = settings.TableName },
            "create-from-dax" => new { sheetName = settings.SheetName, tableName = settings.TableName, daxQuery, targetCell = settings.Range },
            "update-dax" => new { tableName = settings.TableName, daxQuery },
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
        [Description("The action to perform (e.g., list, create, read, append)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--table <NAME>")]
        [Description("Table name")]
        public string? TableName { get; init; }

        [CommandOption("--sheet <NAME>")]
        [Description("Target worksheet name")]
        public string? SheetName { get; init; }

        [CommandOption("--range <ADDRESS>")]
        [Description("Cell range address (e.g., A1:C10)")]
        public string? Range { get; init; }

        [CommandOption("--new-name <NAME>")]
        [Description("New name for rename, or column name for set-column-total")]
        public string? NewName { get; init; }

        [CommandOption("--style <NAME>")]
        [Description("Table style name or total function")]
        public string? Style { get; init; }

        [CommandOption("--has-headers")]
        [Description("Table has header row (default: true)")]
        public bool HasHeaders { get; init; } = true;

        [CommandOption("--csv-data <DATA>")]
        [Description("CSV data to append")]
        public string? CsvData { get; init; }

        [CommandOption("--csv-file <PATH>")]
        [Description("Read CSV data from file instead of command line")]
        public string? CsvFile { get; init; }

        [CommandOption("--visible-only")]
        [Description("Get only visible data (excludes filtered rows)")]
        public bool VisibleOnly { get; init; }

        [CommandOption("--dax-query <QUERY>")]
        [Description("DAX query for create-from-dax or update-dax")]
        public string? DaxQuery { get; init; }

        [CommandOption("--dax-query-file <PATH>")]
        [Description("Read DAX query from file instead of command line")]
        public string? DaxQueryFile { get; init; }
    }
}
