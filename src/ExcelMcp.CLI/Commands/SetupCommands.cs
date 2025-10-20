using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Setup and configuration commands for ExcelCLI - wraps Core commands with CLI formatting
/// </summary>
public class SetupCommands : ISetupCommands
{
    private readonly Core.Commands.SetupCommands _coreCommands = new();
    
    public int EnableVbaTrust(string[] args)
    {
        AnsiConsole.MarkupLine("[cyan]Enabling VBA project access trust...[/]");
        
        var result = _coreCommands.EnableVbaTrust();
        
        if (result.Success)
        {
            // Show which paths were set
            foreach (var path in result.RegistryPathsSet)
            {
                AnsiConsole.MarkupLine($"[green]✓[/] Set VBA trust in: {path}");
            }
            
            AnsiConsole.MarkupLine("[green]✓[/] VBA project access trust has been enabled!");
            
            if (!string.IsNullOrEmpty(result.ManualInstructions))
            {
                AnsiConsole.MarkupLine($"[yellow]Note:[/] {result.ManualInstructions}");
            }
            
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            
            if (!string.IsNullOrEmpty(result.ManualInstructions))
            {
                AnsiConsole.MarkupLine($"[yellow]Manual setup:[/]");
                foreach (var line in result.ManualInstructions.Split('\n'))
                {
                    AnsiConsole.MarkupLine($"  {line}");
                }
            }
            
            return 1;
        }
    }

    public int CheckVbaTrust(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] check-vba-trust <test-file.xlsx>");
            AnsiConsole.MarkupLine("[yellow]Note:[/] Provide a test Excel file to verify VBA access");
            return 1;
        }

        string testFile = args[1];
        
        AnsiConsole.MarkupLine("[cyan]Checking VBA project access trust...[/]");
        
        var result = _coreCommands.CheckVbaTrust(testFile);
        
        if (result.Success && result.IsTrusted)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] VBA project access is [green]TRUSTED[/]");
            AnsiConsole.MarkupLine($"[dim]Found {result.ComponentCount} VBA components in workbook[/]");
            return 0;
        }
        else
        {
            if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage) && !result.ErrorMessage.Contains("not found"))
            {
                // File not found or other error
                AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage.EscapeMarkup()}");
            }
            else
            {
                // Not trusted
                AnsiConsole.MarkupLine($"[red]✗[/] VBA project access is [red]NOT TRUSTED[/]");
                if (!string.IsNullOrEmpty(result.ErrorMessage))
                {
                    AnsiConsole.MarkupLine($"[dim]Error: {result.ErrorMessage.EscapeMarkup()}[/]");
                }
            }
            
            if (!string.IsNullOrEmpty(result.ManualInstructions))
            {
                AnsiConsole.MarkupLine("");
                AnsiConsole.MarkupLine("[yellow]To enable VBA access:[/]");
                AnsiConsole.MarkupLine("1. Run: [cyan]ExcelCLI setup-vba-trust[/]");
                AnsiConsole.MarkupLine("2. Or manually: File → Options → Trust Center → Trust Center Settings → Macro Settings");
                AnsiConsole.MarkupLine("3. Check: 'Trust access to the VBA project object model'");
            }
            
            return 1;
        }
    }
}