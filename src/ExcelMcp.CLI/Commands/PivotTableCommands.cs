using Spectre.Console;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// PivotTable management commands implementation for CLI
/// Wraps Core commands and provides console formatting
/// </summary>
public class PivotTableCommands
{
    private readonly Core.Commands.PivotTable.PivotTableCommands _coreCommands = new();

    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing file path");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] pivot-list <file.xlsx>");
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
            if (result.PivotTables == null || !result.PivotTables.Any())
            {
                AnsiConsole.MarkupLine("[yellow]No PivotTables found in workbook[/]");
                return 0;
            }

            var table = new Table();
            table.AddColumn("PivotTable Name");
            table.AddColumn("Sheet");
            table.AddColumn("Range");
            table.AddColumn("Source Data");
            table.AddColumn("Row Fields");
            table.AddColumn("Column Fields");
            table.AddColumn("Value Fields");

            foreach (var pt in result.PivotTables)
            {
                table.AddRow(
                    pt.Name,
                    pt.SheetName,
                    pt.Range,
                    pt.SourceData,
                    pt.RowFieldCount.ToString(),
                    pt.ColumnFieldCount.ToString(),
                    pt.ValueFieldCount.ToString()
                );
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine($"\n[dim]Found {result.PivotTables.Count} PivotTable(s)[/]");

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int CreateFromRange(string[] args)
    {
        if (args.Length < 7)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] pivot-create-from-range <file.xlsx> <sourceSheet> <sourceRange> <destSheet> <destCell> <pivotTableName>");
            AnsiConsole.MarkupLine("[dim]Example:[/] pivot-create-from-range sales.xlsx Data A1:D100 Analysis A1 SalesPivot");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string sourceSheet = args[2];
        string sourceRange = args[3];
        string destSheet = args[4];
        string destCell = args[5];
        string pivotTableName = args[6];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.CreateFromRangeAsync(batch, sourceSheet, sourceRange, destSheet, destCell, pivotTableName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]Success:[/] Created PivotTable '{result.PivotTableName}'");
            AnsiConsole.MarkupLine($"[dim]Location:[/] {result.SheetName}!{result.Range}");
            AnsiConsole.MarkupLine($"[dim]Source:[/] {result.SourceData} ({result.SourceRowCount} rows)");
            
            if (result.AvailableFields.Any())
            {
                AnsiConsole.MarkupLine($"\n[yellow]Available Fields:[/]");
                foreach (var field in result.AvailableFields)
                {
                    string fieldType = "";
                    if (result.NumericFields.Contains(field)) fieldType = " [dim](numeric)[/]";
                    else if (result.DateFields.Contains(field)) fieldType = " [dim](date)[/]";
                    else if (result.TextFields.Contains(field)) fieldType = " [dim](text)[/]";
                    AnsiConsole.MarkupLine($"  - {field}{fieldType}");
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

    public int AddRowField(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] pivot-add-row-field <file.xlsx> <pivotTableName> <fieldName> [position]");
            AnsiConsole.MarkupLine("[dim]Example:[/] pivot-add-row-field sales.xlsx SalesPivot Region");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string pivotTableName = args[2];
        string fieldName = args[3];
        int? position = args.Length > 4 ? int.Parse(args[4]) : null;

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.AddRowFieldAsync(batch, pivotTableName, fieldName, position);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]Success:[/] Added '{result.FieldName}' to Row area");
            AnsiConsole.MarkupLine($"[dim]Position:[/] {result.Position}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int AddValueField(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] pivot-add-value-field <file.xlsx> <pivotTableName> <fieldName> [function] [customName]");
            AnsiConsole.MarkupLine("[dim]Example:[/] pivot-add-value-field sales.xlsx SalesPivot Sales Sum \"Total Sales\"");
            AnsiConsole.MarkupLine("[dim]Functions:[/] Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string pivotTableName = args[2];
        string fieldName = args[3];
        
        AggregationFunction function = AggregationFunction.Sum;
        if (args.Length > 4)
        {
            if (!Enum.TryParse<AggregationFunction>(args[4], true, out function))
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Invalid function '{args[4]}'");
                AnsiConsole.MarkupLine("[dim]Valid functions:[/] Sum, Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP");
                return 1;
            }
        }
        
        string? customName = args.Length > 5 ? args[5] : null;

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.AddValueFieldAsync(batch, pivotTableName, fieldName, function, customName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]Success:[/] Added '{result.CustomName}' to Values area");
            AnsiConsole.MarkupLine($"[dim]Function:[/] {result.Function}");
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
            AnsiConsole.MarkupLine("[yellow]Usage:[/] pivot-refresh <file.xlsx> <pivotTableName>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string pivotTableName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.RefreshAsync(batch, pivotTableName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]Success:[/] Refreshed PivotTable '{result.PivotTableName}'");
            AnsiConsole.MarkupLine($"[dim]Record Count:[/] {result.SourceRecordCount} (previous: {result.PreviousRecordCount})");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
