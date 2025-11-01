using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Power Query management commands - CLI presentation layer (formats Core results)
/// </summary>
public class PowerQueryCommands : IPowerQueryCommands
{
    private readonly Core.Commands.IPowerQueryCommands _coreCommands;

    public PowerQueryCommands()
    {
        var dataModelCommands = new Core.Commands.DataModelCommands();
        _coreCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
    }

    /// <inheritdoc />
    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-list <file.xlsx>");
            return 1;
        }

        string filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Power Queries in:[/] {Path.GetFileName(filePath)}\n");

        PowerQueryListResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                return await _coreCommands.ListAsync(batch);
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.ErrorMessage?.Contains(".xls") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Note:[/] .xls files don't support Power Query. Use .xlsx or .xlsm");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]This workbook may not have Power Query enabled[/]");
                AnsiConsole.MarkupLine("[dim]Try opening the file in Excel and adding a Power Query first[/]");
            }

            return 1;
        }

        if (result.Queries.Count > 0)
        {
            var table = new Table();
            table.AddColumn("[bold]Query Name[/]");
            table.AddColumn("[bold]Formula (preview)[/]");
            table.AddColumn("[bold]Type[/]");

            foreach (var query in result.Queries.OrderBy(q => q.Name))
            {
                string typeInfo = query.IsConnectionOnly ? "[dim]Connection Only[/]" : "Loaded";

                table.AddRow(
                    $"[cyan]{query.Name.EscapeMarkup()}[/]",
                    $"[dim]{query.FormulaPreview.EscapeMarkup()}[/]",
                    typeInfo
                );
            }

            AnsiConsole.Write(table);
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine($"[bold]Total:[/] {result.Queries.Count} Power Queries");

            // Usage hints
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[dim]Next steps:[/]");
            AnsiConsole.MarkupLine($"[dim]• View query code:[/] [cyan]ExcelCLI pq-view \"{filePath}\" \"QueryName\"[/]");
            AnsiConsole.MarkupLine($"[dim]• Export query:[/] [cyan]ExcelCLI pq-export \"{filePath}\" \"QueryName\" \"output.pq\"[/]");
            AnsiConsole.MarkupLine($"[dim]• Refresh query:[/] [cyan]ExcelCLI pq-refresh \"{filePath}\" \"QueryName\"[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[yellow]No Power Queries found[/]");
            AnsiConsole.MarkupLine("[dim]Create one with:[/] [cyan]ExcelCLI pq-import \"{filePath}\" \"QueryName\" \"code.pq\"[/]");
        }

        return 0;
    }

    /// <inheritdoc />
    public int View(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-view <file.xlsx> <query-name>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];

        PowerQueryViewResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                return await _coreCommands.ViewAsync(batch, queryName);
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.ErrorMessage?.Contains("Did you mean") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Tip:[/] Use [cyan]pq-list[/] to see all available queries");
            }

            return 1;
        }

        AnsiConsole.MarkupLine($"[bold]Power Query:[/] [cyan]{queryName}[/]");
        if (result.IsConnectionOnly)
        {
            AnsiConsole.MarkupLine("[yellow]Type:[/] Connection Only (not loaded to worksheet)");
        }
        else
        {
            AnsiConsole.MarkupLine("[green]Type:[/] Loaded to worksheet");
        }
        AnsiConsole.MarkupLine($"[dim]Characters:[/] {result.CharacterCount}");
        AnsiConsole.WriteLine();

        var panel = new Panel(result.MCode.EscapeMarkup())
        {
            Header = new PanelHeader("Power Query M Code"),
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Blue)
        };
        AnsiConsole.Write(panel);

        return 0;
    }

    /// <inheritdoc />
    public async Task<int> Update(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-update <file.xlsx> <query-name> <mcode-file>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string mCodeFile = args[3];

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await _coreCommands.UpdateAsync(batch, queryName, mCodeFile);
        await batch.SaveAsync();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Updated Power Query '[cyan]{queryName}[/]' from [cyan]{mCodeFile}[/]");
        // Display suggested next actions
        return 0;
    }

    /// <inheritdoc />
    public async Task<int> Export(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-export <file.xlsx> <query-name> [output-file]");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string outputFile = args.Length > 3 ? args[3] : $"{queryName}.pq";

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await _coreCommands.ExportAsync(batch, queryName, outputFile);

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Exported Power Query '[cyan]{queryName}[/]' to [cyan]{outputFile}[/]");

        if (File.Exists(outputFile))
        {
            var fileInfo = new FileInfo(outputFile);
            AnsiConsole.MarkupLine($"[dim]File size: {fileInfo.Length} bytes[/]");
        }

        return 0;
    }

    /// <inheritdoc />
    public async Task<int> Import(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-import <file.xlsx> <query-name> <mcode-file> [--destination worksheet|data-model|both|connection-only]");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string mCodeFile = args[3];
        
        // Parse destination parameter (default: worksheet)
        string loadDestination = "worksheet";
        for (int i = 4; i < args.Length; i++)
        {
            if (args[i].Equals("--destination", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                loadDestination = args[i + 1];
                break;
            }
            // Legacy: --connection-only flag (for backward compatibility)
            else if (args[i].Equals("--connection-only", StringComparison.OrdinalIgnoreCase))
            {
                loadDestination = "connection-only";
                break;
            }
        }

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await _coreCommands.ImportAsync(batch, queryName, mCodeFile, loadDestination: loadDestination);
        await batch.SaveAsync();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.ErrorMessage?.Contains("already exists") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Tip:[/] Use [cyan]pq-update[/] to modify existing queries");
            }

            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Imported Power Query '[cyan]{queryName}[/]' from [cyan]{mCodeFile}[/]");
        // Display suggested next actions
        return 0;
    }

    /// <inheritdoc />
    public int Refresh(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-refresh <file.xlsx> <query-name>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];

        AnsiConsole.MarkupLine($"[bold]Refreshing:[/] [cyan]{queryName}[/]...");

        PowerQueryRefreshResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var refreshResult = await _coreCommands.RefreshAsync(batch, queryName);
                await batch.SaveAsync();
                return refreshResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        if (result.ErrorMessage?.Contains("connection-only") == true)
        {
            AnsiConsole.MarkupLine($"[yellow]Note:[/] {result.ErrorMessage}");
        }
        else
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Refreshed Power Query '[cyan]{queryName}[/]'");
        }
        return 0;
    }

    /// <inheritdoc />
    public int Errors(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-errors <file.xlsx> <query-name>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ErrorsAsync(batch, queryName);
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[bold]Error Status for:[/] [cyan]{queryName}[/]");
        AnsiConsole.MarkupLine(result.MCode.EscapeMarkup());

        return 0;
    }

    /// <inheritdoc />
    public int LoadTo(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-loadto <file.xlsx> <query-name> <sheet-name>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string sheetName = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var loadResult = await _coreCommands.LoadToAsync(batch, queryName, sheetName);
            await batch.SaveAsync();
            return loadResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Loaded Power Query '[cyan]{queryName}[/]' to worksheet '[cyan]{sheetName}[/]'");
        return 0;
    }

    /// <inheritdoc />
    public int Delete(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-delete <file.xlsx> <query-name>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];

        if (!AnsiConsole.Confirm($"Delete Power Query '[cyan]{queryName}[/]'?"))
        {
            AnsiConsole.MarkupLine("[yellow]Cancelled[/]");
            return 1;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var deleteResult = await _coreCommands.DeleteAsync(batch, queryName);
            await batch.SaveAsync();
            return deleteResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Deleted Power Query '[cyan]{queryName}[/]'");
        return 0;
    }

    /// <inheritdoc />
    public int Sources(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-sources <file.xlsx>");
            return 1;
        }

        string filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Excel.CurrentWorkbook() sources in:[/] {Path.GetFileName(filePath)}\n");
        AnsiConsole.MarkupLine("[dim]This shows what tables/ranges Power Query can see[/]\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.SourcesAsync(batch);
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        if (result.Worksheets.Count > 0)
        {
            var table = new Table();
            table.AddColumn("[bold]Name[/]");
            table.AddColumn("[bold]Type[/]");

            // Categorize sources
            var tables = result.Worksheets.Where(w => w.Index <= 1000).ToList();
            var namedRanges = result.Worksheets.Where(w => w.Index > 1000).ToList();

            foreach (var item in tables)
            {
                table.AddRow($"[cyan]{item.Name.EscapeMarkup()}[/]", "Table");
            }

            foreach (var item in namedRanges)
            {
                table.AddRow($"[yellow]{item.Name.EscapeMarkup()}[/]", "Named Range");
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine($"\n[dim]Total: {result.Worksheets.Count} sources[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[yellow]No sources found[/]");
        }

        return 0;
    }

    /// <inheritdoc />
    public int Test(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-test <file.xlsx> <source-name>");
            return 1;
        }

        string filePath = args[1];
        string sourceName = args[2];

        AnsiConsole.MarkupLine($"[bold]Testing source:[/] [cyan]{sourceName}[/]\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.TestAsync(batch, sourceName);
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[yellow]Tip:[/] Use '[cyan]pq-sources[/]' to see all available sources");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Source '[cyan]{sourceName}[/]' exists and can be loaded");

        if (result.ErrorMessage != null)
        {
            AnsiConsole.MarkupLine($"\n[yellow]⚠[/] {result.ErrorMessage}");
        }
        else
        {
            AnsiConsole.MarkupLine($"\n[green]✓[/] Query refreshes successfully");
        }

        AnsiConsole.MarkupLine($"\n[dim]Power Query M code to use:[/]");
        string mCode = $"Excel.CurrentWorkbook(){{{{[Name=\"{sourceName}\"]}}}}[Content]";
        var panel = new Panel(mCode.EscapeMarkup())
        {
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Grey)
        };
        AnsiConsole.Write(panel);

        return 0;
    }

    /// <inheritdoc />
    public int Peek(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-peek <file.xlsx> <source-name>");
            return 1;
        }

        string filePath = args[1];
        string sourceName = args[2];

        AnsiConsole.MarkupLine($"[bold]Preview of:[/] [cyan]{sourceName}[/]\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.PeekAsync(batch, sourceName);
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[yellow]Tip:[/] Use '[cyan]pq-sources[/]' to see all available sources");
            return 1;
        }

        if (result.Data.Count > 0)
        {
            AnsiConsole.MarkupLine($"[green]Named Range Value:[/] {result.Data[0][0]}");
            AnsiConsole.MarkupLine($"[dim]Type: Single cell or range[/]");
        }
        else if (result.ColumnCount > 0)
        {
            AnsiConsole.MarkupLine($"[green]Table found:[/]");
            AnsiConsole.MarkupLine($"  Rows: {result.RowCount}");
            AnsiConsole.MarkupLine($"  Columns: {result.ColumnCount}");

            if (result.Headers.Count > 0)
            {
                string columns = string.Join(", ", result.Headers);
                if (result.ColumnCount > result.Headers.Count)
                {
                    columns += "...";
                }
                AnsiConsole.MarkupLine($"  Columns: {columns}");
            }
        }

        return 0;
    }

    /// <inheritdoc />
    public int Eval(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-eval <file.xlsx> <m-expression>");
            Console.WriteLine("Example: pq-eval Plan.xlsx \"Excel.CurrentWorkbook(){[Name='Growth']}[Content]\"");
            AnsiConsole.MarkupLine("[dim]Purpose:[/] Validates Power Query M syntax and checks if expression can evaluate");
            return 1;
        }

        string filePath = args[1];
        string mExpression = args[2];

        AnsiConsole.MarkupLine($"[bold]Evaluating M expression:[/]\n");
        AnsiConsole.MarkupLine($"[dim]{mExpression.EscapeMarkup()}[/]\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.EvalAsync(batch, mExpression);
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        if (result.ErrorMessage != null)
        {
            AnsiConsole.MarkupLine($"[yellow]⚠[/] Expression syntax is valid but refresh failed");
            AnsiConsole.MarkupLine($"[dim]{result.ErrorMessage.EscapeMarkup()}[/]");
        }
        else
        {
            AnsiConsole.MarkupLine($"[green]✓[/] M expression is valid and can be evaluated");
        }

        return 0;
    }

    /// <summary>
    /// Sets a Power Query to Connection Only mode
    /// </summary>
    public int SetConnectionOnly(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-set-connection-only <file.xlsx> <queryName>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];

        AnsiConsole.MarkupLine($"[bold]Setting '{queryName}' to Connection Only mode...[/]");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SetConnectionOnlyAsync(batch, queryName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Query '{queryName}' is now Connection Only");
        return 0;
    }

    /// <summary>
    /// Sets a Power Query to Load to Table mode
    /// </summary>
    public int SetLoadToTable(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-set-load-to-table <file.xlsx> <queryName> <sheetName>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string sheetName = args[3];

        AnsiConsole.MarkupLine($"[bold]Setting '{queryName}' to Load to Table mode (atomic operation, sheet: {sheetName})...[/]");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SetLoadToTableAsync(batch, queryName, sheetName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        // Check for privacy error (indicated in error message)
        if (!result.Success && result.ErrorMessage?.Contains("privacy level", StringComparison.OrdinalIgnoreCase) == true)
        {
            AnsiConsole.MarkupLine($"[yellow]Privacy Level Required[/]");
            AnsiConsole.MarkupLine($"[dim]{result.ErrorMessage.EscapeMarkup()}[/]");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[dim]Workflow Status: {result.WorkflowStatus}[/]");
            return 1;
        }

        // Display success with verification details
        AnsiConsole.MarkupLine($"[green]✓[/] Query '{queryName}' loaded to worksheet '{sheetName}'");
        AnsiConsole.MarkupLine($"[dim]Rows Loaded: {result.RowsLoaded}[/]");
        AnsiConsole.MarkupLine($"[dim]Workflow Status: {result.WorkflowStatus}[/]");
        return 0;
    }

    /// <summary>
    /// Sets a Power Query to Load to Data Model mode
    /// </summary>
    public int SetLoadToDataModel(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-set-load-to-data-model <file.xlsx> <queryName>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];

        AnsiConsole.MarkupLine($"[bold]Setting '{queryName}' to Load to Data Model mode (atomic operation)...[/]");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SetLoadToDataModelAsync(batch, queryName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        // Check for privacy error (indicated in error message)
        if (!result.Success && result.ErrorMessage?.Contains("privacy level", StringComparison.OrdinalIgnoreCase) == true)
        {
            AnsiConsole.MarkupLine($"[yellow]Privacy Level Required[/]");
            AnsiConsole.MarkupLine($"[dim]{result.ErrorMessage.EscapeMarkup()}[/]");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[dim]Workflow Status: {result.WorkflowStatus}[/]");
            return 1;
        }

        // Display success with verification details
        AnsiConsole.MarkupLine($"[green]✓[/] Query '{queryName}' loaded to Data Model");
        AnsiConsole.MarkupLine($"[dim]Rows Loaded: {result.RowsLoaded}, Tables in Data Model: {result.TablesInDataModel}[/]");
        AnsiConsole.MarkupLine($"[dim]Workflow Status: {result.WorkflowStatus}[/]");
        return 0;
    }

    /// <summary>
    /// Sets a Power Query to Load to Both modes
    /// </summary>
    public int SetLoadToBoth(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-set-load-to-both <file.xlsx> <queryName> <sheetName>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string sheetName = args[3];

        AnsiConsole.MarkupLine($"[bold]Setting '{queryName}' to Load to Both modes (atomic operation, table + data model, sheet: {sheetName})...[/]");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SetLoadToBothAsync(batch, queryName, sheetName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        // Check for privacy error (indicated in error message)
        if (!result.Success && result.ErrorMessage?.Contains("privacy level", StringComparison.OrdinalIgnoreCase) == true)
        {
            AnsiConsole.MarkupLine($"[yellow]Privacy Level Required[/]");
            AnsiConsole.MarkupLine($"[dim]{result.ErrorMessage.EscapeMarkup()}[/]");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[dim]Workflow Status: {result.WorkflowStatus}[/]");
            return 1;
        }

        // Display success with comprehensive verification details
        AnsiConsole.MarkupLine($"[green]✓[/] Query '{queryName}' loaded to both destinations");

        var table = new Table()
            .Border(TableBorder.Rounded)
            .AddColumn("Destination")
            .AddColumn("Status")
            .AddColumn("Rows Loaded");

        table.AddRow(
            "Worksheet Table",
            result.DataLoadedToTable ? "[green]✓ Success[/]" : "[red]✗ Failed[/]",
            result.RowsLoadedToTable.ToString()
        );

        table.AddRow(
            "Data Model",
            result.DataLoadedToModel ? "[green]✓ Success[/]" : "[red]✗ Failed[/]",
            result.RowsLoadedToModel.ToString()
        );

        AnsiConsole.Write(table);

        AnsiConsole.MarkupLine($"[dim]Total Tables in Data Model: {result.TablesInDataModel}[/]");
        AnsiConsole.MarkupLine($"[dim]Workflow Status: {result.WorkflowStatus}[/]");
        return 0;
    }

    /// <summary>
    /// Gets the current load configuration of a Power Query
    /// </summary>
    public int GetLoadConfig(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-get-load-config <file.xlsx> <queryName>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];

        AnsiConsole.MarkupLine($"[bold]Getting load configuration for '{queryName}'...[/]\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetLoadConfigAsync(batch, queryName);
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        var table = new Table()
            .Border(TableBorder.Rounded)
            .AddColumn("Property")
            .AddColumn("Value");

        table.AddRow("Query Name", result.QueryName);
        table.AddRow("Load Mode", result.LoadMode.ToString());
        table.AddRow("Has Connection", result.HasConnection ? "Yes" : "No");
        table.AddRow("Target Sheet", result.TargetSheet ?? "None");
        table.AddRow("Loaded to Data Model", result.IsLoadedToDataModel ? "Yes" : "No");

        AnsiConsole.Write(table);

        // Add helpful information based on load mode
        AnsiConsole.WriteLine();
        switch (result.LoadMode)
        {
            case Core.Models.PowerQueryLoadMode.ConnectionOnly:
                AnsiConsole.MarkupLine("[dim]Connection Only: Query data is not loaded to worksheet or data model[/]");
                break;
            case Core.Models.PowerQueryLoadMode.LoadToTable:
                AnsiConsole.MarkupLine("[dim]Load to Table: Query data is loaded to worksheet[/]");
                break;
            case Core.Models.PowerQueryLoadMode.LoadToDataModel:
                AnsiConsole.MarkupLine("[dim]Load to Data Model: Query data is loaded to PowerPivot data model[/]");
                break;
            case Core.Models.PowerQueryLoadMode.LoadToBoth:
                AnsiConsole.MarkupLine("[dim]Load to Both: Query data is loaded to both worksheet and data model[/]");
                break;
        }

        return 0;
    }
}
