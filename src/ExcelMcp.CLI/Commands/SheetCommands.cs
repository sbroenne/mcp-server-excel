using Spectre.Console;
using Sbroenne.ExcelMcp.Core.Security;

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

        // Validate and normalize CSV file path to prevent path traversal attacks
        try
        {
            csvFile = PathValidator.ValidateExistingFile(csvFile, nameof(csvFile));
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Invalid CSV file path: {ex.Message.EscapeMarkup()}");
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

        // Validate and normalize CSV file path to prevent path traversal attacks
        try
        {
            csvFile = PathValidator.ValidateExistingFile(csvFile, nameof(csvFile));
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Invalid CSV file path: {ex.Message.EscapeMarkup()}");
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

    public int Protect(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-protect <file.xlsx> <sheetName> [password] [options]");
            AnsiConsole.MarkupLine("[dim]Options: --allow-format-cells --allow-format-columns --allow-format-rows");
            AnsiConsole.MarkupLine("         --allow-insert-columns --allow-insert-rows --allow-delete-columns");
            AnsiConsole.MarkupLine("         --allow-delete-rows --allow-sort --allow-filter[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        string? password = args.Length > 3 && !args[3].StartsWith("--") ? args[3] : null;

        // Parse permission flags
        bool allowFormatCells = args.Any(a => a == "--allow-format-cells");
        bool allowFormatColumns = args.Any(a => a == "--allow-format-columns");
        bool allowFormatRows = args.Any(a => a == "--allow-format-rows");
        bool allowInsertColumns = args.Any(a => a == "--allow-insert-columns");
        bool allowInsertRows = args.Any(a => a == "--allow-insert-rows");
        bool allowDeleteColumns = args.Any(a => a == "--allow-delete-columns");
        bool allowDeleteRows = args.Any(a => a == "--allow-delete-rows");
        bool allowSort = args.Any(a => a == "--allow-sort");
        bool allowFilter = args.Any(a => a == "--allow-filter");

        AnsiConsole.MarkupLine($"[bold]Protecting worksheet:[/] '{sheetName.EscapeMarkup()}' in {Path.GetFileName(filePath)}");
        if (!string.IsNullOrEmpty(password))
        {
            AnsiConsole.MarkupLine("[dim]Password protection enabled[/]");
        }

        var result = _coreCommands.Protect(filePath, sheetName, password, 
            allowFormatCells, allowFormatColumns, allowFormatRows,
            allowInsertColumns, allowInsertRows, allowDeleteColumns,
            allowDeleteRows, allowSort, allowFilter);

        if (result.Success)
        {
            AnsiConsole.MarkupLine("[green]✓ Worksheet protected successfully[/]");
            
            // Show enabled permissions
            var permissions = new List<string>();
            if (allowFormatCells) permissions.Add("Format cells");
            if (allowFormatColumns) permissions.Add("Format columns");
            if (allowFormatRows) permissions.Add("Format rows");
            if (allowInsertColumns) permissions.Add("Insert columns");
            if (allowInsertRows) permissions.Add("Insert rows");
            if (allowDeleteColumns) permissions.Add("Delete columns");
            if (allowDeleteRows) permissions.Add("Delete rows");
            if (allowSort) permissions.Add("Sort");
            if (allowFilter) permissions.Add("Filter");

            if (permissions.Count > 0)
            {
                AnsiConsole.MarkupLine("\n[bold]Allowed permissions:[/]");
                foreach (var perm in permissions)
                {
                    AnsiConsole.MarkupLine($"  • {perm}");
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[dim]No user permissions enabled (full protection)[/]");
            }

            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Unprotect(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-unprotect <file.xlsx> <sheetName> [password]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        string? password = args.Length > 3 ? args[3] : null;

        AnsiConsole.MarkupLine($"[bold]Unprotecting worksheet:[/] '{sheetName.EscapeMarkup()}' in {Path.GetFileName(filePath)}");

        var result = _coreCommands.Unprotect(filePath, sheetName, password);

        if (result.Success)
        {
            AnsiConsole.MarkupLine("[green]✓ Worksheet unprotected successfully[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetProtectionStatus(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] sheet-get-protection-status <file.xlsx> <sheetName>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        AnsiConsole.MarkupLine($"[bold]Protection status for:[/] '{sheetName.EscapeMarkup()}' in {Path.GetFileName(filePath)}\n");

        var result = _coreCommands.GetProtectionStatus(filePath, sheetName);

        if (result.Success)
        {
            var table = new Table();
            table.AddColumn("[bold]Property[/]");
            table.AddColumn("[bold]Value[/]");

            table.AddRow("Protected", result.IsProtected ? "[green]Yes[/]" : "[red]No[/]");
            
            if (result.IsProtected)
            {
                table.AddRow("Password Protected", result.HasPassword ? "[yellow]Yes[/]" : "No");
                table.AddRow("Allow Format Cells", result.AllowFormattingCells ? "[green]✓[/]" : "[red]✗[/]");
                table.AddRow("Allow Format Columns", result.AllowFormattingColumns ? "[green]✓[/]" : "[red]✗[/]");
                table.AddRow("Allow Format Rows", result.AllowFormattingRows ? "[green]✓[/]" : "[red]✗[/]");
                table.AddRow("Allow Insert Columns", result.AllowInsertingColumns ? "[green]✓[/]" : "[red]✗[/]");
                table.AddRow("Allow Insert Rows", result.AllowInsertingRows ? "[green]✓[/]" : "[red]✗[/]");
                table.AddRow("Allow Delete Columns", result.AllowDeletingColumns ? "[green]✓[/]" : "[red]✗[/]");
                table.AddRow("Allow Delete Rows", result.AllowDeletingRows ? "[green]✓[/]" : "[red]✗[/]");
                table.AddRow("Allow Sort", result.AllowSorting ? "[green]✓[/]" : "[red]✗[/]");
                table.AddRow("Allow Filter", result.AllowFiltering ? "[green]✓[/]" : "[red]✗[/]");
            }

            AnsiConsole.Write(table);
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
