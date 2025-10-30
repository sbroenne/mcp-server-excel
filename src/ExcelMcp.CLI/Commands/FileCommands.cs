using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// File management commands implementation for CLI
/// Wraps Core commands and provides console formatting
/// </summary>
public class FileCommands : IFileCommands
{
    private readonly Core.Commands.FileCommands _coreCommands = new();

    public int CreateEmpty(string[] args)
    {
        // Validate arguments
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing file path");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] create-empty <file.xlsx|file.xlsm>");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);

        // Check if file already exists and ask for confirmation
        bool overwrite = false;
        if (File.Exists(filePath))
        {
            AnsiConsole.MarkupLine($"[yellow]Warning:[/] File already exists: {filePath}");

            if (!AnsiConsole.Confirm("Do you want to overwrite the existing file?"))
            {
                AnsiConsole.MarkupLine("[dim]Operation cancelled.[/]");
                return 1;
            }
            overwrite = true;
        }

        // Call core command
        var task = Task.Run(async () =>
        {
            return await _coreCommands.CreateEmptyAsync(filePath, overwrite);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success)
        {
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension == ".xlsm")
            {
                AnsiConsole.MarkupLine($"[green]✓[/] Created macro-enabled Excel workbook: [cyan]{Path.GetFileName(filePath)}[/]");
            }
            else
            {
                AnsiConsole.MarkupLine($"[green]✓[/] Created Excel workbook: [cyan]{Path.GetFileName(filePath)}[/]");
            }
            AnsiConsole.MarkupLine($"[dim]Full path: {filePath}[/]");

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

            // Provide helpful tips based on error
            if (result.ErrorMessage?.Contains("extension") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Tip:[/] Use .xlsm for macro-enabled workbooks");
            }

            return 1;
        }
    }
}
