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

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ListAsync(batch);
        });
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

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetAsync(batch, queryTableName);
        });
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

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.DeleteAsync(batch, queryTableName);
            await batch.SaveAsync();
            return result;
        });
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
}
