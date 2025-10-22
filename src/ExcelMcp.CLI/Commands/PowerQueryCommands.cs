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
        _coreCommands = new Core.Commands.PowerQueryCommands();
    }

    /// <summary>
    /// Parses privacy level from command line arguments or environment variable
    /// </summary>
    private static PowerQueryPrivacyLevel? ParsePrivacyLevel(string[] args)
    {
        // Check for --privacy-level parameter
        for (int i = 0; i < args.Length - 1; i++)
        {
            if (args[i] == "--privacy-level" && i + 1 < args.Length)
            {
                if (Enum.TryParse<PowerQueryPrivacyLevel>(args[i + 1], ignoreCase: true, out var level))
                {
                    return level;
                }
            }
        }

        // Check environment variable as fallback
        string? envLevel = Environment.GetEnvironmentVariable("EXCEL_DEFAULT_PRIVACY_LEVEL");
        if (!string.IsNullOrEmpty(envLevel))
        {
            if (Enum.TryParse<PowerQueryPrivacyLevel>(envLevel, ignoreCase: true, out var level))
            {
                return level;
            }
        }

        return null;
    }

    /// <summary>
    /// Displays privacy consent prompt when PowerQueryPrivacyErrorResult is encountered
    /// </summary>
    private static void DisplayPrivacyConsentPrompt(PowerQueryPrivacyErrorResult error)
    {
        AnsiConsole.WriteLine();

        var panel = new Panel(new Markup(
            $"[yellow]Power Query Privacy Level Required[/]\n\n" +
            $"Your query combines data from multiple sources. Excel requires a privacy level to be specified.\n\n" +
            (error.ExistingPrivacyLevels.Count > 0
                ? $"[cyan]Existing queries in this workbook:[/]\n" +
                  string.Join("\n", error.ExistingPrivacyLevels.Select(q => $"  • {q.QueryName}: {q.PrivacyLevel}")) + "\n\n"
                : "") +
            $"[cyan]Recommended:[/] {error.RecommendedPrivacyLevel}\n" +
            $"{error.Explanation}\n\n" +
            $"[dim]To proceed, run the command again with:[/]\n" +
            $"  --privacy-level {error.RecommendedPrivacyLevel}\n\n" +
            $"[dim]Or choose a different level:[/]\n" +
            $"  --privacy-level None          (least secure, ignores privacy)\n" +
            $"  --privacy-level Private       (most secure, prevents combining)\n" +
            $"  --privacy-level Organizational (internal data sources)\n" +
            $"  --privacy-level Public        (public data sources)"
        ));
        panel.Header = new PanelHeader("[yellow]⚠ User Consent Required[/]");
        panel.Border = BoxBorder.Rounded;
        panel.BorderStyle = new Style(Color.Yellow);

        AnsiConsole.Write(panel);
        AnsiConsole.WriteLine();
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

        var result = _coreCommands.List(filePath);

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

        var result = _coreCommands.View(filePath, queryName);

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
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-update <file.xlsx> <query-name> <mcode-file> [--privacy-level <None|Private|Organizational|Public>]");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string mCodeFile = args[3];
        var privacyLevel = ParsePrivacyLevel(args);

        var result = await _coreCommands.Update(filePath, queryName, mCodeFile, privacyLevel);

        // Handle privacy error result
        if (result is PowerQueryPrivacyErrorResult privacyError)
        {
            DisplayPrivacyConsentPrompt(privacyError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Updated Power Query '[cyan]{queryName}[/]' from [cyan]{mCodeFile}[/]");

        // Display workflow hints if available
        if (!string.IsNullOrEmpty(result.WorkflowHint))
        {
            AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
        }

        // Display suggested next actions
        if (result.SuggestedNextActions?.Any() == true)
        {
            AnsiConsole.MarkupLine("[yellow]Suggested next steps:[/]");
            foreach (var action in result.SuggestedNextActions)
            {
                AnsiConsole.MarkupLine($"  [dim]•[/] {action.EscapeMarkup()}");
            }
        }

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

        var result = await _coreCommands.Export(filePath, queryName, outputFile);

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
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-import <file.xlsx> <query-name> <mcode-file> [--privacy-level <None|Private|Organizational|Public>] [--connection-only]");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string mCodeFile = args[3];
        var privacyLevel = ParsePrivacyLevel(args);
        bool loadToWorksheet = !args.Any(a => a.Equals("--connection-only", StringComparison.OrdinalIgnoreCase));

        var result = await _coreCommands.Import(filePath, queryName, mCodeFile, privacyLevel, autoRefresh: true, loadToWorksheet: loadToWorksheet);

        // Handle privacy error result
        if (result is PowerQueryPrivacyErrorResult privacyError)
        {
            DisplayPrivacyConsentPrompt(privacyError);
            return 1;
        }

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

        // Display workflow hints if available
        if (!string.IsNullOrEmpty(result.WorkflowHint))
        {
            AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
        }

        // Display suggested next actions
        if (result.SuggestedNextActions?.Any() == true)
        {
            AnsiConsole.MarkupLine("[yellow]Suggested next steps:[/]");
            foreach (var action in result.SuggestedNextActions)
            {
                AnsiConsole.MarkupLine($"  [dim]•[/] {action.EscapeMarkup()}");
            }
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

        var result = _coreCommands.Refresh(filePath, queryName);

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

        // Display workflow hints if available
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

        var result = _coreCommands.Errors(filePath, queryName);

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

        var result = _coreCommands.LoadTo(filePath, queryName, sheetName);

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

        var result = _coreCommands.Delete(filePath, queryName);

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Deleted Power Query '[cyan]{queryName}[/]'");

        // Display workflow hints if available
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

        var result = _coreCommands.Sources(filePath);

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

        var result = _coreCommands.Test(filePath, sourceName);

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

        var result = _coreCommands.Peek(filePath, sourceName);

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

        var result = _coreCommands.Eval(filePath, mExpression);

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

        var result = _coreCommands.SetConnectionOnly(filePath, queryName);

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Query '{queryName}' is now Connection Only");

        // Display workflow hints if available
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

    /// <summary>
    /// Sets a Power Query to Load to Table mode
    /// </summary>
    public int SetLoadToTable(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-set-load-to-table <file.xlsx> <queryName> <sheetName> [--privacy-level <None|Private|Organizational|Public>]");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string sheetName = args[3];
        var privacyLevel = ParsePrivacyLevel(args);

        AnsiConsole.MarkupLine($"[bold]Setting '{queryName}' to Load to Table mode (sheet: {sheetName})...[/]");

        var result = _coreCommands.SetLoadToTable(filePath, queryName, sheetName, privacyLevel);

        // Handle privacy error result
        if (result is PowerQueryPrivacyErrorResult privacyError)
        {
            DisplayPrivacyConsentPrompt(privacyError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Query '{queryName}' is now loading to worksheet '{sheetName}'");

        // Display workflow hints if available
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

    /// <summary>
    /// Sets a Power Query to Load to Data Model mode
    /// </summary>
    public int SetLoadToDataModel(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-set-load-to-data-model <file.xlsx> <queryName> [--privacy-level <None|Private|Organizational|Public>]");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        var privacyLevel = ParsePrivacyLevel(args);

        AnsiConsole.MarkupLine($"[bold]Setting '{queryName}' to Load to Data Model mode...[/]");

        var result = _coreCommands.SetLoadToDataModel(filePath, queryName, privacyLevel);

        // Handle privacy error result
        if (result is PowerQueryPrivacyErrorResult privacyError)
        {
            DisplayPrivacyConsentPrompt(privacyError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Query '{queryName}' is now loading to Data Model");

        // Display workflow hints if available
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

    /// <summary>
    /// Sets a Power Query to Load to Both modes
    /// </summary>
    public int SetLoadToBoth(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] pq-set-load-to-both <file.xlsx> <queryName> <sheetName> [--privacy-level <None|Private|Organizational|Public>]");
            return 1;
        }

        string filePath = args[1];
        string queryName = args[2];
        string sheetName = args[3];
        var privacyLevel = ParsePrivacyLevel(args);

        AnsiConsole.MarkupLine($"[bold]Setting '{queryName}' to Load to Both modes (table + data model, sheet: {sheetName})...[/]");

        var result = _coreCommands.SetLoadToBoth(filePath, queryName, sheetName, privacyLevel);

        // Handle privacy error result
        if (result is PowerQueryPrivacyErrorResult privacyError)
        {
            DisplayPrivacyConsentPrompt(privacyError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]✗[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Query '{queryName}' is now loading to both worksheet '{sheetName}' and Data Model");

        // Display workflow hints if available
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

        var result = _coreCommands.GetLoadConfig(filePath, queryName);

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
