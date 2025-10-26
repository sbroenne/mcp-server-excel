using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Session;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

public class HyperlinkCommands : IHyperlinkCommands
{
    private readonly Core.Commands.HyperlinkCommands _coreCommands = new();

    public int AddHyperlink(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] hyperlink-add <file.xlsx> <sheet> <cell> <url> [displayText] [tooltip]");
            AnsiConsole.MarkupLine("[dim]Example:[/] hyperlink-add sales.xlsx Sheet1 A1 \"https://example.com\" \"Click here\"");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string sheetName = args[2];
        string cellAddress = args[3];
        string url = args[4];
        string? displayText = args.Length > 5 ? args[5] : null;
        string? tooltip = args.Length > 6 ? args[6] : null;

        // Call core command
        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.AddHyperlinkAsync(batch, sheetName, cellAddress, url, displayText, tooltip);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Added hyperlink to [cyan]{cellAddress}[/] in '{sheetName}'");

            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[dim]Suggested Next Actions:[/]");
                foreach (var action in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  [dim]•[/] {action.EscapeMarkup()}");
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

    public int RemoveHyperlink(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] hyperlink-remove <file.xlsx> <sheet> <cell>");
            AnsiConsole.MarkupLine("[dim]Example:[/] hyperlink-remove sales.xlsx Sheet1 A1");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string sheetName = args[2];
        string cellAddress = args[3];

        // Confirm removal
        if (!AnsiConsole.Confirm($"Remove hyperlink from '{cellAddress}' in '{sheetName}'?"))
        {
            AnsiConsole.MarkupLine("[dim]Operation cancelled.[/]");
            return 1;
        }

        // Call core command
        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.RemoveHyperlinkAsync(batch, sheetName, cellAddress);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Removed hyperlink from [cyan]{cellAddress}[/] in '{sheetName}'");

            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[dim]Suggested Next Actions:[/]");
                foreach (var action in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  [dim]•[/] {action.EscapeMarkup()}");
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

    public int ListHyperlinks(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] hyperlink-list <file.xlsx> <sheet>");
            AnsiConsole.MarkupLine("[dim]Example:[/] hyperlink-list sales.xlsx Sheet1");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string sheetName = args[2];

        // Call core command
        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ListHyperlinksAsync(batch, sheetName);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success)
        {
            if (result.Hyperlinks == null || !result.Hyperlinks.Any())
            {
                AnsiConsole.MarkupLine("[yellow]No hyperlinks found in sheet '{0}'[/]", sheetName.EscapeMarkup());
                return 0;
            }

            var table = new Table();
            table.AddColumn("Cell");
            table.AddColumn("URL");
            table.AddColumn("Display Text");
            table.AddColumn("Tooltip");

            foreach (var link in result.Hyperlinks)
            {
                table.AddRow(
                    link.CellAddress.EscapeMarkup(),
                    link.Address?.EscapeMarkup() ?? "[dim]N/A[/]",
                    link.DisplayText?.EscapeMarkup() ?? "[dim]N/A[/]",
                    link.ScreenTip?.EscapeMarkup() ?? "[dim]N/A[/]"
                );
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine($"\n[dim]Found {result.Hyperlinks.Count} hyperlink(s) in '{sheetName}'[/]");

            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[dim]Suggested Next Actions:[/]");
                foreach (var action in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  [dim]•[/] {action.EscapeMarkup()}");
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

    public int GetHyperlink(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Error:[/] Missing required arguments");
            AnsiConsole.MarkupLine("[yellow]Usage:[/] hyperlink-get <file.xlsx> <sheet> <cell>");
            AnsiConsole.MarkupLine("[dim]Example:[/] hyperlink-get sales.xlsx Sheet1 A1");
            return 1;
        }

        string filePath = Path.GetFullPath(args[1]);
        string sheetName = args[2];
        string cellAddress = args[3];

        // Call core command
        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetHyperlinkAsync(batch, sheetName, cellAddress);
        });
        var result = task.GetAwaiter().GetResult();

        // Format and display result
        if (result.Success && result.Hyperlink != null)
        {
            var panel = new Panel($"""
                [bold]Cell:[/] [cyan]{result.Hyperlink.CellAddress}[/]
                [bold]URL:[/] {result.Hyperlink.Address?.EscapeMarkup() ?? "[dim]N/A[/]"}
                [bold]Display Text:[/] {result.Hyperlink.DisplayText?.EscapeMarkup() ?? "[dim]N/A[/]"}
                [bold]Tooltip:[/] {result.Hyperlink.ScreenTip?.EscapeMarkup() ?? "[dim]N/A[/]"}
                """)
            {
                Header = new PanelHeader($"Hyperlink Information - {sheetName}"),
                Border = BoxBorder.Rounded
            };

            AnsiConsole.Write(panel);

            if (!string.IsNullOrEmpty(result.WorkflowHint))
            {
                AnsiConsole.MarkupLine($"\n[dim]{result.WorkflowHint.EscapeMarkup()}[/]");
            }

            if (result.SuggestedNextActions != null && result.SuggestedNextActions.Any())
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[dim]Suggested Next Actions:[/]");
                foreach (var action in result.SuggestedNextActions)
                {
                    AnsiConsole.MarkupLine($"  [dim]•[/] {action.EscapeMarkup()}");
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

    // Core interface implementations (not used in CLI - batch-of-one pattern)
    public async Task<Core.Models.OperationResult> AddHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress, string url, string? displayText = null, string? tooltip = null)
        => await _coreCommands.AddHyperlinkAsync(batch, sheetName, cellAddress, url, displayText, tooltip);

    public async Task<Core.Models.OperationResult> RemoveHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress)
        => await _coreCommands.RemoveHyperlinkAsync(batch, sheetName, cellAddress);

    public async Task<Core.Models.HyperlinkListResult> ListHyperlinksAsync(IExcelBatch batch, string sheetName)
        => await _coreCommands.ListHyperlinksAsync(batch, sheetName);

    public async Task<Core.Models.HyperlinkInfoResult> GetHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress)
        => await _coreCommands.GetHyperlinkAsync(batch, sheetName, cellAddress);
}
