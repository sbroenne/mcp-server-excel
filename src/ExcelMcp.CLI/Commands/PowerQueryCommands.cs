using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Power Query management commands - CLI presentation layer (formats Core results)
/// </summary>
public class PowerQueryCommands : IPowerQueryCommands
{
    private readonly Core.Commands.PowerQueryCommands _coreCommands;

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
            AnsiConsole.MarkupLine("[dim]Create one with:[/] [cyan]ExcelCLI pq-create \"{filePath}\" \"QueryName\" \"code.pq\"[/]");
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
            return await _coreCommands.ListExcelSourcesAsync(batch);
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
    [Obsolete("pq-test has been removed. Use table-info or parameter-get instead.")]
    public int Test(string[] args)
    {
        AnsiConsole.MarkupLine("[red]Command Removed:[/] pq-test has been deprecated.\n");
        AnsiConsole.MarkupLine("[yellow]This command was confusing because it tested Excel sources (tables/ranges), not Power Query queries.[/]\n");

        AnsiConsole.MarkupLine("[bold]Use instead:[/]");
        AnsiConsole.MarkupLine("  [cyan]table-info[/] file.xlsx TableName          Check if table exists (returns info)");
        AnsiConsole.MarkupLine("  [cyan]parameter-get[/] file.xlsx RangeName      Check if named range exists (returns value)");
        AnsiConsole.MarkupLine("  [cyan]pq-sources[/] file.xlsx                   List all available sources\n");

        AnsiConsole.MarkupLine("[dim]Note: If the operation succeeds, the source exists. If it fails, it doesn't.[/]");
        return 1;
    }

    [Obsolete("pq-peek has been removed. Use table-info or parameter-get instead.")]
    public int Peek(string[] args)
    {
        AnsiConsole.MarkupLine("[red]Command Removed:[/] pq-peek has been deprecated.\n");
        AnsiConsole.MarkupLine("[yellow]This command was confusing because it peeked at Excel sources (tables/ranges), not Power Query queries.[/]\n");

        AnsiConsole.MarkupLine("[bold]Use instead:[/]");
        AnsiConsole.MarkupLine("  [cyan]table-info[/] file.xlsx TableName          Preview table (includes row/column count, headers)");
        AnsiConsole.MarkupLine("  [cyan]parameter-get[/] file.xlsx RangeName      Get named range value");
        AnsiConsole.MarkupLine("  [cyan]pq-view[/] file.xlsx QueryName           View Power Query M code\n");

        AnsiConsole.MarkupLine("[dim]The new commands are more intuitive: table commands for tables, parameter commands for ranges.[/]");
        return 1;
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
            case PowerQueryLoadMode.ConnectionOnly:
                AnsiConsole.MarkupLine("[dim]Connection Only: Query data is not loaded to worksheet or data model[/]");
                break;
            case PowerQueryLoadMode.LoadToTable:
                AnsiConsole.MarkupLine("[dim]Load to Table: Query data is loaded to worksheet[/]");
                break;
            case PowerQueryLoadMode.LoadToDataModel:
                AnsiConsole.MarkupLine("[dim]Load to Data Model: Query data is loaded to PowerPivot data model[/]");
                break;
            case PowerQueryLoadMode.LoadToBoth:
                AnsiConsole.MarkupLine("[dim]Load to Both: Query data is loaded to both worksheet and data model[/]");
                break;
        }

        return 0;
    }

    // ========================================
    // Atomic Operations
    // ========================================

    /// <summary>
    /// Creates a new Power Query from M code file with atomic import + load
    /// </summary>
    public async Task<int> Create(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-create <file.xlsx> <query-name> <mcode-file> [--destination worksheet|data-model|both|connection-only] [--target-sheet SheetName]");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string mCodeFile = args[3];

        // Parse optional parameters
        PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable; // Default: worksheet
        string? targetSheet = null;

        for (int i = 4; i < args.Length; i++)
        {
            if (args[i].Equals("--destination", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                loadMode = args[i + 1].ToLowerInvariant() switch
                {
                    "worksheet" => PowerQueryLoadMode.LoadToTable,
                    "data-model" => PowerQueryLoadMode.LoadToDataModel,
                    "both" => PowerQueryLoadMode.LoadToBoth,
                    "connection-only" => PowerQueryLoadMode.ConnectionOnly,
                    _ => PowerQueryLoadMode.LoadToTable
                };
                i++;
            }
            else if (args[i].Equals("--target-sheet", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
            {
                targetSheet = args[i + 1];
                i++;
            }
        }

        AnsiConsole.MarkupLine($"[bold]Creating Power Query '[cyan]{queryName}[/]'...[/]");

        PowerQueryCreateResult result;
        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            result = await _coreCommands.CreateAsync(batch, queryName, mCodeFile, loadMode, targetSheet);
            await batch.SaveAsync();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Created Power Query '[cyan]{queryName}[/]'");
        AnsiConsole.MarkupLine($"[dim]Load Destination:[/] {result.LoadDestination}");
        if (result.WorksheetName != null)
        {
            AnsiConsole.MarkupLine($"[dim]Worksheet:[/] {result.WorksheetName}");
        }
        if (result.DataLoaded)
        {
            AnsiConsole.MarkupLine($"[dim]Rows Loaded:[/] {result.RowsLoaded}");
        }

        // Workflow guidance - CLI layer responsibility
        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[dim]Suggested next actions:[/]");
        AnsiConsole.MarkupLine($"  • Use 'pq-refresh {filePath} {queryName}' to update data");
        if (result.LoadDestination is PowerQueryLoadMode.LoadToDataModel or PowerQueryLoadMode.LoadToBoth)
        {
            AnsiConsole.MarkupLine($"  • Use 'dm-create-measure' to add DAX calculations");
        }
        AnsiConsole.MarkupLine($"  • Use 'pq-view {filePath} {queryName}' to inspect M code");

        return 0;
    }

    /// <summary>
    /// Updates only the M code of a Power Query (no refresh)
    /// </summary>
    public async Task<int> UpdateMCode(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-update-mcode <file.xlsx> <query-name> <mcode-file>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string mCodeFile = args[3];

        AnsiConsole.MarkupLine($"[bold]Updating M code for '[cyan]{queryName}[/]' (no refresh)...[/]");

        OperationResult result;
        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            result = await _coreCommands.UpdateMCodeAsync(batch, queryName, mCodeFile);
            await batch.SaveAsync();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Updated M code for '[cyan]{queryName}[/]'");
        AnsiConsole.MarkupLine("[yellow]Note:[/] Query data not refreshed. Use [cyan]pq-refresh[/] to refresh data.");

        return 0;
    }

    /// <summary>
    /// Converts a Power Query to connection-only (unloads data from worksheet/data model)
    /// </summary>
    public async Task<int> Unload(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-unload <file.xlsx> <query-name>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];

        AnsiConsole.MarkupLine($"[bold]Converting '[cyan]{queryName}[/]' to connection-only...[/]");

        OperationResult result;
        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            result = await _coreCommands.UnloadAsync(batch, queryName);
            await batch.SaveAsync();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Converted '[cyan]{queryName}[/]' to connection-only");
        AnsiConsole.MarkupLine("[dim]Query is now connection-only (data removed from worksheet/data model)[/]");

        return 0;
    }

    /// <summary>
    /// Updates M code AND refreshes data in one atomic operation
    /// </summary>
    public async Task<int> UpdateAndRefresh(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-update-and-refresh <file.xlsx> <query-name> <mcode-file>");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string mCodeFile = args[3];

        AnsiConsole.MarkupLine($"[bold]Updating and refreshing '[cyan]{queryName}[/]'...[/]");

        OperationResult result;
        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            result = await _coreCommands.UpdateAndRefreshAsync(batch, queryName, mCodeFile);
            await batch.SaveAsync();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Updated and refreshed '[cyan]{queryName}[/]'");
        AnsiConsole.MarkupLine("[dim]M code updated and data refreshed in one operation[/]");

        return 0;
    }

    /// <summary>
    /// Refreshes all Power Queries in the workbook in one operation
    /// </summary>
    public async Task<int> RefreshAll(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-refresh-all <file.xlsx>");
            return 1;
        }

        string filePath = args[1];

        AnsiConsole.MarkupLine("[bold]Refreshing all Power Queries...[/]");

        OperationResult result;
        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            result = await _coreCommands.RefreshAllAsync(batch);
            await batch.SaveAsync();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine("[green]✓[/] All Power Queries refreshed successfully");

        return 0;
    }
}
