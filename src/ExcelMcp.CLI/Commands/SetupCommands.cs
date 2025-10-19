using Spectre.Console;
using Microsoft.Win32;
using System;
using static ExcelMcp.ExcelHelper;

namespace ExcelMcp.Commands;

/// <summary>
/// Setup and configuration commands for ExcelCLI
/// </summary>
public class SetupCommands : ISetupCommands
{
    /// <summary>
    /// Enable VBA project access trust in Excel registry
    /// </summary>
    public int EnableVbaTrust(string[] args)
    {
        try
        {
            AnsiConsole.MarkupLine("[cyan]Enabling VBA project access trust...[/]");
            
            // Try different Office versions and architectures
            string[] registryPaths = {
                @"SOFTWARE\Microsoft\Office\16.0\Excel\Security",  // Office 2019/2021/365
                @"SOFTWARE\Microsoft\Office\15.0\Excel\Security",  // Office 2013
                @"SOFTWARE\Microsoft\Office\14.0\Excel\Security",  // Office 2010
                @"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Excel\Security",  // 32-bit on 64-bit
                @"SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Excel\Security",
                @"SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Excel\Security"
            };

            bool successfullySet = false;
            
            foreach (string path in registryPaths)
            {
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(path))
                    {
                        if (key != null)
                        {
                            // Set AccessVBOM = 1 to trust VBA project access
                            key.SetValue("AccessVBOM", 1, RegistryValueKind.DWord);
                            AnsiConsole.MarkupLine($"[green]✓[/] Set VBA trust in: {path}");
                            successfullySet = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[dim]Skipped {path}: {ex.Message.EscapeMarkup()}[/]");
                }
            }

            if (successfullySet)
            {
                AnsiConsole.MarkupLine("[green]✓[/] VBA project access trust has been enabled!");
                AnsiConsole.MarkupLine("[yellow]Note:[/] You may need to restart Excel for changes to take effect.");
                return 0;
            }
            else
            {
                AnsiConsole.MarkupLine("[red]Error:[/] Could not find Excel registry keys to modify.");
                AnsiConsole.MarkupLine("[yellow]Manual setup:[/] File → Options → Trust Center → Trust Center Settings → Macro Settings");
                AnsiConsole.MarkupLine("[yellow]Manual setup:[/] Check 'Trust access to the VBA project object model'");
                return 1;
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }

    /// <summary>
    /// Check current VBA trust status
    /// </summary>
    public int CheckVbaTrust(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] check-vba-trust <test-file.xlsx>");
            AnsiConsole.MarkupLine("[yellow]Note:[/] Provide a test Excel file to verify VBA access");
            return 1;
        }

        string testFile = args[1];
        if (!File.Exists(testFile))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Test file not found: {testFile}");
            return 1;
        }

        try
        {
            AnsiConsole.MarkupLine("[cyan]Checking VBA project access trust...[/]");
            
            int result = WithExcel(testFile, false, (excel, workbook) =>
            {
                try
                {
                    dynamic vbProject = workbook.VBProject;
                    int componentCount = vbProject.VBComponents.Count;
                    
                    AnsiConsole.MarkupLine($"[green]✓[/] VBA project access is [green]TRUSTED[/]");
                    AnsiConsole.MarkupLine($"[dim]Found {componentCount} VBA components in workbook[/]");
                    return 0;
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]✗[/] VBA project access is [red]NOT TRUSTED[/]");
                    AnsiConsole.MarkupLine($"[dim]Error: {ex.Message.EscapeMarkup()}[/]");
                    
                    AnsiConsole.MarkupLine("");
                    AnsiConsole.MarkupLine("[yellow]To enable VBA access:[/]");
                    AnsiConsole.MarkupLine("1. Run: [cyan]ExcelCLI setup-vba-trust[/]");
                    AnsiConsole.MarkupLine("2. Or manually: File → Options → Trust Center → Trust Center Settings → Macro Settings");
                    AnsiConsole.MarkupLine("3. Check: 'Trust access to the VBA project object model'");
                    
                    return 1;
                }
            });
            
            return result;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error testing VBA access:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }
}