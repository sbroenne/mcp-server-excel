using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Worksheet management commands - wraps Core with CLI formatting
/// </summary>
public class SheetCommands : ISheetCommands
{
    private readonly Core.Commands.SheetCommands _coreCommands = new();

    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-list <file.xlsx>");
            return 1;
        }

        var filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Worksheets in:[/] {Path.GetFileName(filePath)}\n");

        var result = _coreCommands.List(filePath);

        if (result.Success)
        {
            if (result.Worksheets.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Index[/]");
                table.AddColumn("[bold]Worksheet Name[/]");

                foreach (var sheet in result.Worksheets)
                {
                    table.AddRow(sheet.Index.ToString(), sheet.Name.EscapeMarkup());
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"\n[dim]Found {result.Worksheets.Count} worksheet(s)[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No worksheets found[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Read(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-read <file.xlsx> <sheet-name> <range>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var range = args[3];

        var result = _coreCommands.Read(filePath, sheetName, range);

        if (result.Success)
        {
            foreach (var row in result.Data)
            {
                var values = row.Select(v => v?.ToString() ?? "").ToArray();
                Console.WriteLine(string.Join(",", values));
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public async Task<int> Write(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-write <file.xlsx> <sheet-name> <csv-file>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var csvFile = args[3];

        if (!File.Exists(csvFile))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] CSV file not found: {csvFile}");
            return 1;
        }

        var csvData = await File.ReadAllTextAsync(csvFile);
        var result = _coreCommands.Write(filePath, sheetName, csvData);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Wrote data to {sheetName}");
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
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-create <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        var result = _coreCommands.Create(filePath, sheetName);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Created worksheet '{sheetName.EscapeMarkup()}'");

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
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Rename(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-rename <file.xlsx> <old-name> <new-name>");
            return 1;
        }

        var filePath = args[1];
        var oldName = args[2];
        var newName = args[3];

        var result = _coreCommands.Rename(filePath, oldName, newName);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Renamed '{oldName.EscapeMarkup()}' to '{newName.EscapeMarkup()}'");

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
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Copy(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-copy <file.xlsx> <source-name> <target-name>");
            return 1;
        }

        var filePath = args[1];
        var sourceName = args[2];
        var targetName = args[3];

        var result = _coreCommands.Copy(filePath, sourceName, targetName);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Copied '{sourceName.EscapeMarkup()}' to '{targetName.EscapeMarkup()}'");

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
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-delete <file.xlsx> <sheet-name>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        var result = _coreCommands.Delete(filePath, sheetName);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted worksheet '{sheetName.EscapeMarkup()}'");

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
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Clear(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-clear <file.xlsx> <sheet-name> <range>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var range = args[3];

        var result = _coreCommands.Clear(filePath, sheetName, range);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Cleared range {range.EscapeMarkup()} in {sheetName.EscapeMarkup()}");

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
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Append(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-append <file.xlsx> <sheet-name> <csv-file>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var csvFile = args[3];

        if (!File.Exists(csvFile))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] CSV file not found: {csvFile}");
            return 1;
        }

        var csvData = File.ReadAllText(csvFile);
        var result = _coreCommands.Append(filePath, sheetName, csvData);

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Appended data to {sheetName.EscapeMarkup()}");

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
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
