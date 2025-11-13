using System.Diagnostics.CodeAnalysis;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Table management commands implementation for CLI
/// Wraps Core commands and provides console formatting
/// </summary>
[SuppressMessage("Performance", "CA1861:Avoid constant arrays as arguments", Justification = "Simple workflow arrays in CLI formatting")]
public class CliTableCommands : ITableCommands
{
    private readonly TableCommands _coreCommands = new();

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

        // Call core command with batch
        var task = Task.Run(async () =>
        {
            return await _coreCommands.ListAsync(filePath);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success)
        {
            if (result.Tables == null || (result.Tables.Count == 0))
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
                    t.RowCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    t.ColumnCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
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

        // Call core command directly with filePath
        var task = Task.Run(async () =>
        {
            return await _coreCommands.CreateAsync(filePath, sheetName, tableName, range, hasHeaders, tableStyle);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Created table: [cyan]{tableName}[/]");
            AnsiConsole.MarkupLine($"[dim]Sheet: {sheetName}, Range: {range}[/]");
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
        var task = Task.Run(async () =>
        {
            return await _coreCommands.RenameAsync(filePath, tableName, newName);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Renamed table: [cyan]{tableName}[/] → [cyan]{newName}[/]");
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
        var task = Task.Run(async () =>
        {
            return await _coreCommands.DeleteAsync(filePath, tableName);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted table: [cyan]{tableName}[/]");
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
        var task = Task.Run(async () =>
        {
            return await _coreCommands.GetAsync(filePath, tableName);
        });
        var result = task.GetAwaiter().GetResult();

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

            if ((result.Table.Columns.Count > 0))
            {
                AnsiConsole.MarkupLine("\n[bold]Columns:[/]");
                var columnTable = new Table();
                columnTable.AddColumn("#");
                columnTable.AddColumn("Column Name");

                for (int i = 0; i < result.Table.Columns.Count; i++)
                {
                    columnTable.AddRow((i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), result.Table.Columns[i]);
                }

                AnsiConsole.Write(columnTable);
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

        var task = Task.Run(async () =>
        {
            return await _coreCommands.ResizeAsync(filePath, tableName, newRange);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Resized table: [cyan]{tableName}[/] to {newRange}");
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

        var task = Task.Run(async () =>
        {
            return await _coreCommands.ToggleTotalsAsync(filePath, tableName, showTotals);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Totals row {(showTotals ? "enabled" : "disabled")} for: [cyan]{tableName}[/]");
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

        var task = Task.Run(async () =>
        {
            return await _coreCommands.SetColumnTotalAsync(filePath, tableName, columnName, totalFunction);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set [[cyan]{columnName}[/]] total to {totalFunction}");
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

        // Parse CSV to List<List<object?>>
        var rows = ParseCsvToRows(csvData);

        var task = Task.Run(async () =>
        {
            return await _coreCommands.AppendAsync(filePath, tableName, rows);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Appended rows to table: [cyan]{tableName}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    /// <summary>
    /// Parse CSV data into List of List of objects for table operations.
    /// Simple CSV parser - assumes comma delimiter, handles quoted strings.
    /// </summary>
    private static List<List<object?>> ParseCsvToRows(string csvData)
    {
        var rows = new List<List<object?>>();
        var lines = csvData.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

        foreach (var line in lines)
        {
            var values = line.Split(',');
            var row = values.Select(value =>
            {
                var trimmed = value.Trim().Trim('"');
                return string.IsNullOrEmpty(trimmed) ? null : (object?)trimmed;
            }).ToList();

            rows.Add(row);
        }

        return rows;
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

        var task = Task.Run(async () =>
        {
            return await _coreCommands.SetStyleAsync(filePath, tableName, tableStyle);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Changed style of [cyan]{tableName}[/] to {tableStyle}");
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

        var task = Task.Run(async () =>
        {
            return await _coreCommands.AddToDataModelAsync(filePath, tableName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Added [cyan]{tableName}[/] to Power Pivot Data Model");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === FILTER OPERATIONS ===

    public int ApplyFilter(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-apply-filter <file.xlsx> <tableName> <columnName> <criteria>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-apply-filter sales.xlsx SalesTable Amount \">100\"");
            AnsiConsole.MarkupLine("[dim]Criteria:[/] >value, <value, =value, >=value, <=value, <>value");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string columnName = args[3];
        string criteria = args[4];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ApplyFilterAsync(batch, tableName, columnName, criteria);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Applied filter to column [cyan]{columnName}[/] in table [cyan]{tableName}[/]");
            AnsiConsole.MarkupLine($"[dim]Filter criteria: {criteria.EscapeMarkup()}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ApplyFilterValues(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-apply-filter-values <file.xlsx> <tableName> <columnName> <value1,value2,...>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-apply-filter-values sales.xlsx SalesTable Region \"North,South,East\"");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string columnName = args[3];
        string valuesStr = args[4];

        // Parse comma-separated values
        var filterValues = valuesStr.Split(',').Select(v => v.Trim()).ToList();

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ApplyFilterAsync(batch, tableName, columnName, filterValues);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Applied filter to column [cyan]{columnName}[/] in table [cyan]{tableName}[/]");
            AnsiConsole.MarkupLine($"[dim]Filter values: {string.Join(", ", filterValues)}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ClearFilters(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-clear-filters <file.xlsx> <tableName>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-clear-filters sales.xlsx SalesTable");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ClearFiltersAsync(batch, tableName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Cleared all filters from table [cyan]{tableName}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetFilters(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-get-filters <file.xlsx> <tableName>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-get-filters sales.xlsx SalesTable");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetFiltersAsync(batch, tableName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[cyan]Table:[/] {result.TableName}");
            AnsiConsole.MarkupLine($"[cyan]Has Active Filters:[/] {(result.HasActiveFilters ? "Yes" : "No")}");

            if (result.ColumnFilters != null && (result.ColumnFilters.Count > 0))
            {
                var table = new Table();
                table.AddColumn("Column");
                table.AddColumn("Filtered");
                table.AddColumn("Criteria");
                table.AddColumn("Filter Values");

                foreach (var filter in result.ColumnFilters)
                {
                    table.AddRow(
                        filter.ColumnName,
                        filter.IsFiltered ? "Yes" : "No",
                        filter.Criteria ?? "-",
                        filter.FilterValues != null && (filter.FilterValues.Count > 0)
                            ? string.Join(", ", filter.FilterValues)
                            : "-"
                    );
                }

                AnsiConsole.Write(table);
            }
            else
            {
                AnsiConsole.MarkupLine("[dim]No column filters found[/]");
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === COLUMN OPERATIONS ===

    public int AddColumn(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-add-column <file.xlsx> <tableName> <columnName> [position]");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-add-column sales.xlsx SalesTable NewColumn");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-add-column sales.xlsx SalesTable NewColumn 2");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string columnName = args[3];
        int? position = args.Length > 4 && int.TryParse(args[4], out int pos) ? pos : null;

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.AddColumnAsync(batch, tableName, columnName, position);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Added column [cyan]{columnName}[/] to table [cyan]{tableName}[/]");
            if (position.HasValue)
            {
                AnsiConsole.MarkupLine($"[dim]Position: {position.Value}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int RemoveColumn(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-remove-column <file.xlsx> <tableName> <columnName>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-remove-column sales.xlsx SalesTable OldColumn");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string columnName = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.RemoveColumnAsync(batch, tableName, columnName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Removed column [cyan]{columnName}[/] from table [cyan]{tableName}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int RenameColumn(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-rename-column <file.xlsx> <tableName> <oldColumnName> <newColumnName>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-rename-column sales.xlsx SalesTable OldName NewName");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string oldColumnName = args[3];
        string newColumnName = args[4];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.RenameColumnAsync(batch, tableName, oldColumnName, newColumnName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Renamed column from [cyan]{oldColumnName}[/] to [cyan]{newColumnName}[/] in table [cyan]{tableName}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    #region Phase 2 Commands

    public int GetStructuredReference(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-get-structured-reference <file.xlsx> <tableName> <region> [[columnName]]");
            AnsiConsole.MarkupLine("[dim]Regions:[/] All, Data, Headers, Totals, ThisRow");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-get-structured-reference sales.xlsx SalesTable Data");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-get-structured-reference sales.xlsx SalesTable Data Amount");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string regionStr = args[3];
        string? columnName = args.Length > 4 ? args[4] : null;

        // Parse region
        if (!Enum.TryParse<TableRegion>(regionStr, ignoreCase: true, out var region))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Invalid region '{regionStr}'");
            AnsiConsole.MarkupLine("[yellow]Valid regions:[/] All, Data, Headers, Totals, ThisRow");
            return 1;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetStructuredReferenceAsync(batch, tableName, region, columnName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            var table = new Table();
            table.AddColumn("Property");
            table.AddColumn("Value");

            table.AddRow("Table Name", result.TableName);
            table.AddRow("Region", result.Region.ToString());
            if (!string.IsNullOrEmpty(result.ColumnName))
            {
                table.AddRow("Column Name", result.ColumnName);
            }
            table.AddRow("Structured Reference", $"[cyan]{result.StructuredReference}[/]");
            table.AddRow("Range Address", $"[dim]{result.RangeAddress}[/]");
            table.AddRow("Row Count", result.RowCount.ToString(System.Globalization.CultureInfo.InvariantCulture));
            table.AddRow("Column Count", result.ColumnCount.ToString(System.Globalization.CultureInfo.InvariantCulture));

            AnsiConsole.Write(table);

            AnsiConsole.MarkupLine($"\n[dim]Use this structured reference in formulas or with RangeCommands[/]");

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Sort(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-sort <file.xlsx> <tableName> <columnName> [[asc|desc]]");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-sort sales.xlsx SalesTable Amount desc");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-sort sales.xlsx SalesTable Region asc");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string columnName = args[3];
        bool ascending = args.Length > 4 ? args[4].Equals("asc", StringComparison.OrdinalIgnoreCase) : true;

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SortAsync(batch, tableName, columnName, ascending);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            string direction = ascending ? "ascending" : "descending";
            AnsiConsole.MarkupLine($"[green]✓[/] Sorted table [cyan]{tableName}[/] by [cyan]{columnName}[/] ({direction})");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SortMulti(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-sort-multi <file.xlsx> <tableName> <column1:asc> <column2:desc> [[column3:asc]]");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-sort-multi sales.xlsx SalesTable Region:asc Amount:desc");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-sort-multi sales.xlsx SalesTable Year:desc Quarter:desc Amount:desc");
            AnsiConsole.MarkupLine("[dim]Note:[/] Excel supports max 3 sort levels");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];

        // Parse sort columns (format: "columnName:asc" or "columnName:desc")
        var sortColumns = new List<TableSortColumn>();
        for (int i = 3; i < args.Length && i < 6; i++) // Max 3 levels (args 3, 4, 5)
        {
            string[] parts = args[i].Split(':');
            if (parts.Length != 2)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Invalid sort column format '{args[i]}'");
                AnsiConsole.MarkupLine("[yellow]Expected format:[/] columnName:asc or columnName:desc");
                return 1;
            }

            string columnName = parts[0];
            bool ascending = parts[1].Equals("asc", StringComparison.OrdinalIgnoreCase);

            sortColumns.Add(new TableSortColumn
            {
                ColumnName = columnName,
                Ascending = ascending
            });
        }

        if (sortColumns.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] No sort columns specified");
            return 1;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SortAsync(batch, tableName, sortColumns);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Sorted table [cyan]{tableName}[/] by {sortColumns.Count} level(s):");
            for (int i = 0; i < sortColumns.Count; i++)
            {
                var col = sortColumns[i];
                string direction = col.Ascending ? "ascending" : "descending";
                AnsiConsole.MarkupLine($"  [dim]{i + 1}.[/] [cyan]{col.ColumnName}[/] ({direction})");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === NUMBER FORMAT OPERATIONS ===

    public int GetColumnNumberFormat(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-get-column-format <file.xlsx> <tableName> <columnName>");
            AnsiConsole.MarkupLine("[dim]Example:[/] table-get-column-format sales.xlsx SalesTable Amount");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string columnName = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetColumnNumberFormatAsync(batch, tableName, columnName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[cyan]Table:[/] {tableName}");
            AnsiConsole.MarkupLine($"[cyan]Column:[/] {columnName}");

            // Display the first format (table columns typically have consistent formatting)
            string format = result.Formats != null && result.Formats.Count > 0 && result.Formats[0].Count > 0
                ? result.Formats[0][0]
                : "(none)";

            AnsiConsole.MarkupLine($"[cyan]Number Format:[/] {format.EscapeMarkup()}");

            if (result.RowCount > 1)
            {
                AnsiConsole.MarkupLine($"[dim]Note: Format applies to {result.RowCount} rows[/]");
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetColumnNumberFormat(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] table-set-column-format <file.xlsx> <tableName> <columnName> <formatCode>");
            AnsiConsole.MarkupLine("[dim]Examples:[/]");
            AnsiConsole.MarkupLine("  table-set-column-format sales.xlsx SalesTable Amount \"$#,##0.00\"");
            AnsiConsole.MarkupLine("  table-set-column-format sales.xlsx SalesTable Percent \"0.00%\"");
            AnsiConsole.MarkupLine("  table-set-column-format sales.xlsx SalesTable Date \"m/d/yyyy\"");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string tableName = args[2];
        string columnName = args[3];
        string formatCode = args[4];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SetColumnNumberFormatAsync(batch, tableName, columnName, formatCode);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set number format for column [cyan]{columnName}[/] in table [cyan]{tableName}[/]");
            AnsiConsole.MarkupLine($"[dim]Format: {formatCode}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    #endregion
}

