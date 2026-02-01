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

        if (!ActionValidator.TryNormalizeAction<TableAction>(settings.Action, out var action, out _))
        {
            // Try TableColumnAction if TableAction doesn't match
            var validActions = ActionValidator.GetValidActions<TableAction>()
                .Concat(ActionValidator.GetValidActions<TableColumnAction>())
                .ToArray();

            if (!ActionValidator.TryNormalizeAction(settings.Action, validActions, out action, out var errorMessage))
            {
                AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
                return 1;
            }
        }
        var command = $"table.{action}";

        // Build args based on action
        // Note: property names must match daemon's Args classes (e.g., TableCreateArgs)
        var csvData = ResolveFileOrValue(settings.CsvData, settings.CsvFile);
        var daxQuery = ResolveFileOrValue(settings.DaxQuery, settings.DaxQueryFile);
        object? args = action switch
        {
            // TableAction
            "list" => null,
            "create" => new { sheetName = settings.SheetName, tableName = settings.TableName, range = settings.Range, hasHeaders = settings.HasHeaders, tableStyle = settings.Style },
            "read" => new { tableName = settings.TableName },
            "rename" => new { tableName = settings.TableName, newName = settings.NewName },
            "delete" => new { tableName = settings.TableName },
            "resize" => new { tableName = settings.TableName, newRange = settings.Range },
            "set-style" => new { tableName = settings.TableName, tableStyle = settings.Style },
            "toggle-totals" => new { tableName = settings.TableName, showTotals = settings.ShowTotals },
            "set-column-total" => new { tableName = settings.TableName, columnName = settings.ColumnName, totalFunction = settings.TotalFunction },
            "append" => new { tableName = settings.TableName, csvData },
            "get-data" => new { tableName = settings.TableName, visibleOnly = settings.VisibleOnly },
            "add-to-datamodel" => new { tableName = settings.TableName },
            "create-from-dax" => new { sheetName = settings.SheetName, tableName = settings.TableName, daxQuery, targetCell = settings.Range },
            "update-dax" => new { tableName = settings.TableName, daxQuery },
            "get-dax" => new { tableName = settings.TableName },

            // TableColumnAction
            "apply-filter" => new { tableName = settings.TableName, columnName = settings.ColumnName, criteria = settings.Criteria },
            "apply-filter-values" => new { tableName = settings.TableName, columnName = settings.ColumnName, values = ParseStringList(settings.FilterValues) },
            "clear-filters" => new { tableName = settings.TableName },
            "get-filters" => new { tableName = settings.TableName },
            "add-column" => new { tableName = settings.TableName, columnName = settings.ColumnName, position = settings.ColumnPosition },
            "remove-column" => new { tableName = settings.TableName, columnName = settings.ColumnName },
            "rename-column" => new { tableName = settings.TableName, oldName = settings.ColumnName, newName = settings.NewName },
            "get-structured-reference" => new { tableName = settings.TableName, region = settings.Region, columnName = settings.ColumnName },
            "sort" => new { tableName = settings.TableName, columnName = settings.ColumnName, ascending = settings.Ascending },
            "sort-multi" => new { tableName = settings.TableName, sortColumnsJson = settings.SortColumnsJson },
            "get-column-number-format" => new { tableName = settings.TableName, columnName = settings.ColumnName },
            "set-column-number-format" => new { tableName = settings.TableName, columnName = settings.ColumnName, formatCode = settings.FormatCode },

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

    private static List<string>? ParseStringList(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        try
        {
            return JsonSerializer.Deserialize<List<string>>(input, DaemonProtocol.JsonOptions);
        }
        catch
        {
            return [.. input.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)];
        }
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform (e.g., list, create, read, append)")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--table|--name|--table-name <NAME>")]
        [Description("Table name")]
        public string? TableName { get; init; }

        [CommandOption("--sheet <NAME>")]
        [Description("Target worksheet name")]
        public string? SheetName { get; init; }

        [CommandOption("--range <ADDRESS>")]
        [Description("Cell range address (e.g., A1:C10)")]
        public string? Range { get; init; }

        [CommandOption("--new-name <NAME>")]
        [Description("New name for rename operations")]
        public string? NewName { get; init; }

        [CommandOption("--column <NAME>")]
        [Description("Column name for column operations")]
        public string? ColumnName { get; init; }

        [CommandOption("--style <NAME>")]
        [Description("Table style name")]
        public string? Style { get; init; }

        [CommandOption("--has-headers")]
        [Description("Table has header row (default: true)")]
        public bool HasHeaders { get; init; } = true;

        [CommandOption("--show-totals")]
        [Description("Show totals row")]
        public bool ShowTotals { get; init; }

        [CommandOption("--total-function <FUNCTION>")]
        [Description("Total function (Sum, Count, Average, etc.)")]
        public string? TotalFunction { get; init; }

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

        // TableColumnAction settings
        [CommandOption("--criteria <CRITERIA>")]
        [Description("Filter criteria expression")]
        public string? Criteria { get; init; }

        [CommandOption("--filter-values <VALUES>")]
        [Description("Comma-separated or JSON array of filter values")]
        public string? FilterValues { get; init; }

        [CommandOption("--column-position <NUMBER>")]
        [Description("Column position (0-based index)")]
        public int? ColumnPosition { get; init; }

        [CommandOption("--region <REGION>")]
        [Description("Structured reference region (All, Data, Headers, Totals)")]
        public string? Region { get; init; }

        [CommandOption("--ascending")]
        [Description("Sort ascending (default: true)")]
        public bool Ascending { get; init; } = true;

        [CommandOption("--sort-columns <JSON>")]
        [Description("Sort columns JSON for multi-column sort")]
        public string? SortColumnsJson { get; init; }

        [CommandOption("--format-code <CODE>")]
        [Description("Number format code (e.g., #,##0.00)")]
        public string? FormatCode { get; init; }
    }
}
