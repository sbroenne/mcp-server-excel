using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// File and session management commands implementation for CLI
/// Wraps Core commands and provides console formatting
/// </summary>
public class FileCommands : IFileCommands
{
    private readonly Core.Commands.FileCommands _coreCommands = new();
    private readonly BatchCommands _batchCommands = new();

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

    /// <summary>
    /// Open a session for a workbook - forwards to BatchCommands
    /// </summary>
    public int Open(string[] args) => _batchCommands.Open(args);

    /// <summary>
    /// Save and close a session - forwards to BatchCommands
    /// </summary>
    public int Save(string[] args) => _batchCommands.Save(args);

    /// <summary>
    /// List active sessions - forwards to BatchCommands
    /// </summary>
    public int ListSessions(string[] args) => _batchCommands.List(args);
}
