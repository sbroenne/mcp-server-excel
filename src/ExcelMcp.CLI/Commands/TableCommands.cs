using Spectre.Console;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Table management commands implementation for CLI
/// Wraps Core commands and provides console formatting
/// </summary>
public class TableCommands : ITableCommands
{
    private readonly Core.Commands.TableCommands _coreCommands = new();

    public int List(string[] args)
    {
        // Validate arguments
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing file path");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-list <file.xlsx>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);

        // Call core command
        var result = _coreCommands.List(filePath);

        // Format and display result
        if (result.Success)
        {
            if (result.Tables == null || !result.Tables.Any())
            {
                AnsiConsole.MarkupLine("[yellow]No tables found in workbook[/]");
                return 0;
            }

            var table = new Table();
            table.AddColumn("Table Name");
            table.AddColumn("Sheet");
            table.AddColumn("Range");
            table.AddColumn("Rows");
            table.AddColumn("Columns");
            table.AddColumn("Headers");
            table.AddColumn("Totals");

            foreach (var t in result.Tables)
            {
                table.AddRow(
                    t.Name,
                    t.SheetName,
                    t.Range,
                    t.RowCount.ToString(),
                    t.ColumnCount.ToString(),
                    t.HasHeaders ? "Yes" : "No",
                    t.ShowTotals ? "Yes" : "No"
                );
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine($"\n[dim]Found {result.Tables.Count} table(s)[/]");

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Create(string[] args)
    {
        // Validate arguments
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-create <file.xlsx> <sheetName> <tableName> <range> [hasHeaders] [tableStyle]");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-create sales.xlsx Data SalesTable A1:E100 true TableStyleMedium2");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string sheetName = args[2];
        string tableName = args[3];
        string range = args[4];
        bool hasHeaders = args.Length > 5 ? bool.Parse(args[5]) : true;
        string? tableStyle = args.Length > 6 ? args[6] : null;

        // Call core command
        var result = _coreCommands.Create(filePath, sheetName, tableName, range, hasHeaders, tableStyle);

        // Format and display result
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Created table: [cyan]{tableName}[/]");
            AnsiConsole.MarkupLine($"[dim]Sheet: {sheetName}, Range: {range}[/]");

            // Display workflow hints
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggested Next Actions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Rename(string[] args)
    {
        // Validate arguments
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-rename <file.xlsx> <tableName> <newName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string newName = args[3];

        // Call core command
        var result = _coreCommands.Rename(filePath, tableName, newName);

        // Format and display result
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Renamed table: [cyan]{tableName}[/] → [cyan]{newName}[/]");

            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggested Next Actions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

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
        // Validate arguments
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-delete <file.xlsx> <tableName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];

        // Confirm deletion
        if (!AnsiConsole.Confirm($"Delete table '{tableName}'? (Data will remain as a regular range)"))
        {
            AnsiConsole.MarkupLine("[dim]Operation cancelled.[/]");
            return 1;
        }

        // Call core command
        var result = _coreCommands.Delete(filePath, tableName);

        // Format and display result
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted table: [cyan]{tableName}[/]");

            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggested Next Actions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Info(string[] args)
    {
        // Validate arguments
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-info <file.xlsx> <tableName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];

        // Call core command
        var result = _coreCommands.GetInfo(filePath, tableName);

        // Format and display result
        if (result.Success && result.Table != null)
        {
            var panel = new Panel($"""
                [bold]Table:[/] [cyan]{result.Table.Name}[/]
                [bold]Sheet:[/] {result.Table.SheetName}
                [bold]Range:[/] {result.Table.Range}
                [bold]Rows:[/] {result.Table.RowCount}
                [bold]Columns:[/] {result.Table.ColumnCount}
                [bold]Has Headers:[/] {(result.Table.HasHeaders ? "Yes" : "No")}
                [bold]Show Totals:[/] {(result.Table.ShowTotals ? "Yes" : "No")}
                [bold]Table Style:[/] {result.Table.TableStyle ?? "(none)"}
                """)
            {
                Header = new PanelHeader($"[bold]Table Information[/]"),
                Border = BoxBorder.Rounded
            };

            AnsiConsole.Write(panel);

            if (result.Table.Columns.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Columns:[/]");
                var columnTable = new Table();
                columnTable.AddColumn("#");
                columnTable.AddColumn("Column Name");

                for (int i = 0; i < result.Table.Columns.Count; i++)
                {
                    columnTable.AddRow((i + 1).ToString(), result.Table.Columns[i]);
                }

                AnsiConsole.Write(columnTable);
            }

            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggested Next Actions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Resize(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-resize <file.xlsx> <tableName> <newRange>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-resize sales.xlsx SalesTable A1:E150");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string newRange = args[3];

        var result = _coreCommands.Resize(filePath, tableName, newRange);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Resized table: [cyan]{tableName}[/] to {newRange}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ToggleTotals(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-toggle-totals <file.xlsx> <tableName> <true|false>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        bool showTotals = bool.Parse(args[3]);

        var result = _coreCommands.ToggleTotals(filePath, tableName, showTotals);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Totals row {(showTotals ? "enabled" : "disabled")} for: [cyan]{tableName}[/]");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetColumnTotal(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-set-column-total <file.xlsx> <tableName> <columnName> <function>");
            AnsiConsole.MarkupLine("[dim]Functions:[/] sum, average, count, countnums, max, min, stddev, var, none");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string columnName = args[3];
        string totalFunction = args[4];

        var result = _coreCommands.SetColumnTotal(filePath, tableName, columnName, totalFunction);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set [[cyan]{columnName}[/]] total to {totalFunction}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ReadData(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-read <file.xlsx> <tableName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];

        var result = _coreCommands.ReadData(filePath, tableName);

        if (result.Success)
        {
            if (result.Data == null || !result.Data.Any())
            {
                AnsiConsole.MarkupLine("[yellow]Table is empty[/]");
                return 0;
            }

            var table = new Table();
            
            // Add headers
            foreach (var header in result.Headers)
            {
                table.AddColumn(header);
            }

            // Add data rows
            foreach (var row in result.Data)
            {
                table.AddRow(row.Select(cell => cell?.ToString() ?? "").ToArray());
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine($"\n[dim]{result.RowCount} rows, {result.ColumnCount} columns[/]");

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int AppendRows(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-append <file.xlsx> <tableName> <csvData>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-append sales.xlsx SalesTable \"Product1,100,5.99\\nProduct2,200,3.49\"");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string csvData = args[3];

        var result = _coreCommands.AppendRows(filePath, tableName, csvData);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Appended rows to table: [cyan]{tableName}[/]");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetStyle(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-set-style <file.xlsx> <tableName> <style>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-set-style sales.xlsx SalesTable TableStyleMedium2");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string tableStyle = args[3];

        var result = _coreCommands.SetStyle(filePath, tableName, tableStyle);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Changed style of [cyan]{tableName}[/] to {tableStyle}");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int AddToDataModel(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-add-to-datamodel <file.xlsx> <tableName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];

        var result = _coreCommands.AddToDataModel(filePath, tableName);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Added [cyan]{tableName}[/] to Power Pivot Data Model");
            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }
            
            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.MarkupLine("\n[bold]Suggested Next Actions:[/]");
                foreach (var suggestion in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  • {suggestion.EscapeMarkup()}");
                }
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
