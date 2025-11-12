using Sbroenne.ExcelMcp.ComInterop.Session;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// QueryTable management commands implementation for CLI
/// Wraps Core commands and provides console formatting
/// </summary>
public class QueryTableCommands
{
    private readonly Core.Commands.QueryTable.QueryTableCommands _coreCommands = new();

    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing file path");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] querytable-list <file.xlsx>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);

        var task = Task.Run(async () => await _coreCommands.ListAsync(filePath));
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.QueryTables == null || result.QueryTables.Count == 0)
            {
                AnsiConsole.MarkupLine("[yellow]No QueryTables found in workbook[/]");
                return 0;
            }

            var table = new Table();
            table.AddColumn("QueryTable Name");
            table.AddColumn("Sheet");
            table.AddColumn("Range");
            table.AddColumn("Rows");
            table.AddColumn("Columns");
            table.AddColumn("Last Refresh");

            foreach (var qt in result.QueryTables)
            {
                table.AddRow(
                    qt.Name,
                    qt.WorksheetName,
                    qt.Range,
                    qt.RowCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    qt.ColumnCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    qt.LastRefresh?.ToString("g", System.Globalization.CultureInfo.InvariantCulture) ?? "Never"
                );
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine($"\n[dim]Found {result.QueryTables.Count} QueryTable(s)[/]");

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Get(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] querytable-get <file.xlsx> <queryTableName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string queryTableName = args[2];

        var task = Task.Run(async () => await _coreCommands.GetAsync(filePath, queryTableName));
        var result = task.GetAwaiter().GetResult();

        if (result.Success && result.QueryTable != null)
        {
            var qt = result.QueryTable;
            AnsiConsole.MarkupLine($"[bold]QueryTable:[/] [cyan]{qt.Name}[/]");
            AnsiConsole.MarkupLine($"[bold]Sheet:[/] {qt.WorksheetName}");
            AnsiConsole.MarkupLine($"[bold]Range:[/] {qt.Range}");
            AnsiConsole.MarkupLine($"[bold]Rows:[/] {qt.RowCount}");
            AnsiConsole.MarkupLine($"[bold]Columns:[/] {qt.ColumnCount}");
            AnsiConsole.MarkupLine($"[bold]Background Query:[/] {qt.BackgroundQuery}");
            AnsiConsole.MarkupLine($"[bold]Refresh on Open:[/] {qt.RefreshOnFileOpen}");
            AnsiConsole.MarkupLine($"[bold]Preserve Formatting:[/] {qt.PreserveFormatting}");
            AnsiConsole.MarkupLine($"[bold]Last Refresh:[/] {qt.LastRefresh?.ToString("g", System.Globalization.CultureInfo.InvariantCulture) ?? "Never"}");

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Refresh(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] querytable-refresh <file.xlsx> <queryTableName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string queryTableName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.RefreshAsync(batch, queryTableName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Refreshed QueryTable: [cyan]{queryTableName}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int RefreshAll(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing file path");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] querytable-refresh-all <file.xlsx>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.RefreshAllAsync(batch);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            var count = result.OperationContext?.ContainsKey("RefreshedCount") == true
                ? result.OperationContext["RefreshedCount"]
                : "unknown";
            AnsiConsole.MarkupLine($"[green]✓[/] Refreshed {count} QueryTable(s)");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Delete(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] querytable-delete <file.xlsx> <queryTableName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string queryTableName = args[2];

        var task = Task.Run(async () => await _coreCommands.DeleteAsync(filePath, queryTableName));
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted QueryTable: [cyan]{queryTableName}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int CreateFromConnection(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] querytable-create-from-connection <file.xlsx> <sheetName> <queryTableName> <connectionName> [range] [--background] [--refresh-on-open] [--preserve-formatting]");
            AnsiConsole.MarkupLine("[dim]Example:[/] querytable-create-from-connection data.xlsx Sheet1 MyQT MyConnection A1 --background");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string sheetName = args[2];
        string queryTableName = args[3];
        string connectionName = args[4];
        string range = args.Length > 5 && !args[5].StartsWith("--", StringComparison.Ordinal) ? args[5] : "A1";

        var options = new Core.Models.QueryTableCreateOptions
        {
            BackgroundQuery = args.Contains("--background", StringComparer.OrdinalIgnoreCase),
            RefreshOnFileOpen = args.Contains("--refresh-on-open", StringComparer.OrdinalIgnoreCase),
            PreserveFormatting = args.Contains("--preserve-formatting", StringComparer.OrdinalIgnoreCase) || !args.Contains("--no-preserve-formatting", StringComparer.OrdinalIgnoreCase),
            RefreshImmediately = !args.Contains("--no-refresh", StringComparer.OrdinalIgnoreCase)
        };

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.CreateFromConnectionAsync(batch, sheetName, queryTableName, connectionName, range, options);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Created QueryTable [cyan]{queryTableName}[/] from connection [cyan]{connectionName}[/] on sheet [cyan]{sheetName}[/] at range [cyan]{range}[/]");
            if (options.RefreshImmediately)
            {
                AnsiConsole.MarkupLine("[dim]QueryTable refreshed immediately[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int CreateFromQuery(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] querytable-create-from-query <file.xlsx> <sheetName> <queryTableName> <queryName> [range] [--background] [--refresh-on-open] [--preserve-formatting]");
            AnsiConsole.MarkupLine("[dim]Example:[/] querytable-create-from-query data.xlsx Sheet1 MyQT MyQuery A1 --background");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string sheetName = args[2];
        string queryTableName = args[3];
        string queryName = args[4];
        string range = args.Length > 5 && !args[5].StartsWith("--", StringComparison.Ordinal) ? args[5] : "A1";

        var options = new Core.Models.QueryTableCreateOptions
        {
            BackgroundQuery = args.Contains("--background", StringComparer.OrdinalIgnoreCase),
            RefreshOnFileOpen = args.Contains("--refresh-on-open", StringComparer.OrdinalIgnoreCase),
            PreserveFormatting = args.Contains("--preserve-formatting", StringComparer.OrdinalIgnoreCase) || !args.Contains("--no-preserve-formatting", StringComparer.OrdinalIgnoreCase),
            RefreshImmediately = !args.Contains("--no-refresh", StringComparer.OrdinalIgnoreCase)
        };

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.CreateFromQueryAsync(batch, sheetName, queryTableName, queryName, range, options);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Created QueryTable [cyan]{queryTableName}[/] from Power Query [cyan]{queryName}[/] on sheet [cyan]{sheetName}[/] at range [cyan]{range}[/]");
            if (options.RefreshImmediately)
            {
                AnsiConsole.MarkupLine("[dim]QueryTable refreshed immediately[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int UpdateProperties(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] querytable-update-properties <file.xlsx> <queryTableName> [--background=true|false] [--refresh-on-open=true|false] [--preserve-formatting=true|false]");
            AnsiConsole.MarkupLine("[dim]Example:[/] querytable-update-properties data.xlsx MyQT --background=true --refresh-on-open=false");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string queryTableName = args[2];

        var options = new Core.Models.QueryTableUpdateOptions
        {
            BackgroundQuery = GetBoolOption(args, "--background"),
            RefreshOnFileOpen = GetBoolOption(args, "--refresh-on-open"),
            SavePassword = GetBoolOption(args, "--save-password"),
            PreserveColumnInfo = GetBoolOption(args, "--preserve-column-info"),
            PreserveFormatting = GetBoolOption(args, "--preserve-formatting"),
            AdjustColumnWidth = GetBoolOption(args, "--adjust-column-width")
        };

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.UpdatePropertiesAsync(batch, queryTableName, options);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Updated properties for QueryTable: [cyan]{queryTableName}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    private static bool? GetBoolOption(string[] args, string optionName)
    {
        var option = args.FirstOrDefault(a => a.StartsWith($"{optionName}=", StringComparison.OrdinalIgnoreCase));
        if (option == null) return null;

        var value = option.Split('=')[1];
        return bool.TryParse(value, out var result) ? result : null;
    }
}
