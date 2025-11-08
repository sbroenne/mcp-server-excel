using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// VBA script management commands - CLI presentation layer (formats Core results)
/// </summary>
public class VbaCommands : IVbaCommands
{
    private readonly Core.Commands.VbaCommands _coreCommands;

    public VbaCommands()
    {
        _coreCommands = new Core.Commands.VbaCommands();
    }

    /// <summary>
    /// Displays VBA trust guidance when VbaTrustRequiredResult is encountered
    /// </summary>
    private static void DisplayVbaTrustGuidance(VbaTrustRequiredResult trustError)
    {
        AnsiConsole.WriteLine();

        var panel = new Panel(new Markup(
            $"[yellow]VBA Trust Access Required[/]\n\n" +
            $"{trustError.Explanation}\n\n" +
            $"[cyan]How to enable VBA trust:[/]\n" +
            string.Join("\n", trustError.SetupInstructions.Select((s, i) => $"  {i + 1}. {s}")) + "\n\n" +
            $"[dim]This is a one-time setup. After enabling, VBA operations will work.[/]\n\n" +
            $"[cyan]📖 More information:[/]\n" +
            $"[link]{trustError.DocumentationUrl}[/]"
        ));
        panel.Header = new PanelHeader("[yellow]⚠ Setup Required[/]");
        panel.Border = BoxBorder.Rounded;
        panel.BorderStyle = new Style(Color.Yellow);

        AnsiConsole.Write(panel);

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[dim]After enabling VBA trust in Excel, run this command again.[/]");
    }

    /// <inheritdoc />
    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-list <file.xlsm>");
            return 1;
        }

        string filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]VBA Scripts in:[/] {Path.GetFileName(filePath)}\n");

        VbaListResult result;
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

            if (result.ErrorMessage?.Contains("macro-enabled") == true)
            {
                AnsiConsole.MarkupLine($"[yellow]Current file:[/] {Path.GetFileName(filePath)} ({Path.GetExtension(filePath)})");
                AnsiConsole.MarkupLine($"[yellow]Solutions:[/]");
                AnsiConsole.MarkupLine($"  • Create new .xlsm file: [cyan]ExcelCLI create-empty \"file.xlsm\"[/]");
                AnsiConsole.MarkupLine($"  • Save existing file as .xlsm in Excel");
            }
            else if (result.ErrorMessage?.Contains("not enabled") == true || result.ErrorMessage?.Contains("not trusted") == true)
            {
                AnsiConsole.WriteLine();
                AnsiConsole.MarkupLine("[yellow]VBA trust access is required to list VBA modules.[/]");
                AnsiConsole.MarkupLine("[dim]Enable it manually in Excel:[/] File → Options → Trust Center → Trust Center Settings → Macro Settings");
                AnsiConsole.MarkupLine("[dim]Check '✓ Trust access to the VBA project object model'[/]");
            }

            return 1;
        }

        if (result.Scripts.Count > 0)
        {
            var table = new Table();
            table.AddColumn("[bold]Module Name[/]");
            table.AddColumn("[bold]Type[/]");
            table.AddColumn("[bold]Procedures[/]");

            foreach (var script in result.Scripts.OrderBy(s => s.Name))
            {
                string procedures = script.Procedures.Count > 0
                    ? string.Join(", ", script.Procedures.Take(5)) + (script.Procedures.Count > 5 ? "..." : "")
                    : "[dim](no procedures)[/]";

                table.AddRow(
                    $"[cyan]{script.Name.EscapeMarkup()}[/]",
                    script.Type.EscapeMarkup(),
                    procedures.EscapeMarkup()
                );
            }

            AnsiConsole.Write(table);
            AnsiConsole.MarkupLine($"\n[dim]Total: {result.Scripts.Count} script(s)[/]");

            // Usage hints
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[dim]Next steps:[/]");
            AnsiConsole.MarkupLine($"[dim]• Export script:[/] [cyan]ExcelCLI script-export \"{filePath}\" \"ModuleName\" \"output.vba\"[/]");
            AnsiConsole.MarkupLine($"[dim]• Run procedure:[/] [cyan]ExcelCLI script-run \"{filePath}\" \"ModuleName.ProcedureName\"[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[yellow]No VBA scripts found[/]");
            AnsiConsole.MarkupLine("[dim]Import one with:[/] [cyan]ExcelCLI script-import \"{filePath}\" \"ModuleName\" \"code.vba\"[/]");
        }

        return 0;
    }

    /// <inheritdoc />
    public int View(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-view <file.xlsm> <module-name>");
            AnsiConsole.MarkupLine("\n[bold]Examples:[/]");
            AnsiConsole.MarkupLine("  script-view Report.xlsm DataProcessor");
            return 1;
        }

        string filePath = args[1];
        string moduleName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ViewAsync(batch, moduleName);
        });
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            // Check if it's a VBA trust issue
            if (result.ErrorMessage?.Contains("VBA trust access is not enabled") == true)
            {
                var trustError = new VbaTrustRequiredResult
                {
                    Success = false,
                    ErrorMessage = result.ErrorMessage,
                    IsTrustEnabled = false,
                    Explanation = "VBA operations require 'Trust access to the VBA project object model' to be enabled in Excel settings."
                };
                DisplayVbaTrustGuidance(trustError);
                return 1;
            }

            AnsiConsole.MarkupLine($"[red]✗ Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        // Display module information
        var panel = new Panel($"[bold]Module:[/] {result.ModuleName.EscapeMarkup()}\n" +
                             $"[bold]Type:[/] {result.ModuleType.EscapeMarkup()}\n" +
                             $"[bold]Lines:[/] {result.LineCount}\n" +
                             $"[bold]Procedures:[/] {result.Procedures.Count}\n\n" +
                             $"[bold]Code:[/]\n{result.Code.EscapeMarkup()}")
        {
            Header = new PanelHeader("VBA Module Code"),
            Border = BoxBorder.Rounded
        };
        AnsiConsole.Write(panel);

        if ((result.Procedures.Count > 0))
        {
            AnsiConsole.MarkupLine("\n[bold]Procedures Found:[/]");
            foreach (var proc in result.Procedures)
            {
                AnsiConsole.MarkupLine($"  • {proc.EscapeMarkup()}");
            }
        }

        // Display suggested next actions
        return 0;
    }

    /// <inheritdoc />
    public int Export(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-export <file.xlsm> <module-name> [output-file]");
            return 1;
        }

        string filePath = args[1];
        string moduleName = args[2];
        string outputFile = args.Length > 3 ? args[3] : $"{moduleName}.vba";

        ResultBase result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                return await _coreCommands.ExportAsync(batch, moduleName, outputFile);
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        // Handle VBA trust guidance
        if (result is VbaTrustRequiredResult trustError)
        {
            DisplayVbaTrustGuidance(trustError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.ErrorMessage?.Contains("not found") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Tip:[/] Use [cyan]script-list[/] to see available modules");
            }

            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Exported VBA module '[cyan]{moduleName}[/]' to [cyan]{outputFile}[/]");

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
            AnsiConsole.MarkupLine("[red]Usage:[/] script-import <file.xlsm> <module-name> <vba-file>");
            return 1;
        }

        string filePath = args[1];
        string moduleName = args[2];
        string vbaFile = args[3];

        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await _coreCommands.ImportAsync(batch, moduleName, vbaFile);
        await batch.SaveAsync();

        // Handle VBA trust guidance
        if (result is VbaTrustRequiredResult trustError)
        {
            DisplayVbaTrustGuidance(trustError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.ErrorMessage?.Contains("already exists") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Tip:[/] Use [cyan]script-update[/] to modify existing modules");
            }

            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Imported VBA module '[cyan]{moduleName}[/]' from [cyan]{vbaFile}[/]");
        return 0;
    }

    /// <inheritdoc />
    public async Task<int> Update(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-update <file.xlsm> <module-name> <vba-file>");
            return 1;
        }

        string filePath = args[1];
        string moduleName = args[2];
        string vbaFile = args[3];

        ResultBase result;
        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            result = await _coreCommands.UpdateAsync(batch, moduleName, vbaFile);
            await batch.SaveAsync();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        // Handle VBA trust guidance
        if (result is VbaTrustRequiredResult trustError)
        {
            DisplayVbaTrustGuidance(trustError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");

            if (result.ErrorMessage?.Contains("not found") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Tip:[/] Use [cyan]script-import[/] to create new modules");
            }

            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Updated VBA module '[cyan]{moduleName}[/]' from [cyan]{vbaFile}[/]");
        return 0;
    }

    /// <inheritdoc />
    public int Run(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-run <file.xlsm> <Module.Procedure> [param1] [param2] ...");
            AnsiConsole.MarkupLine("[dim]Example:[/] script-run data.xlsm \"Module1.ProcessData\" \"Sheet1\" \"A1:D100\"");
            return 1;
        }

        string filePath = args[1];
        string procedureName = args[2];
        string[] parameters = args.Skip(3).ToArray();

        AnsiConsole.MarkupLine($"[bold]Running VBA procedure:[/] [cyan]{procedureName}[/]");
        if (parameters.Length > 0)
        {
            AnsiConsole.MarkupLine($"[dim]Parameters:[/] {string.Join(", ", parameters.Select(p => $"\"{p}\""))}");
        }
        AnsiConsole.WriteLine();

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var runResult = await _coreCommands.RunAsync(batch, procedureName, null, parameters);
            await batch.SaveAsync();
            return runResult;
        });
        var result = task.GetAwaiter().GetResult();

        // Handle VBA trust guidance
        if (result is VbaTrustRequiredResult trustError)
        {
            DisplayVbaTrustGuidance(trustError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] VBA procedure '[cyan]{procedureName}[/]' executed successfully");
        return 0;
    }

    /// <inheritdoc />
    public int Delete(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] script-delete <file.xlsm> <module-name>");
            return 1;
        }

        string filePath = args[1];
        string moduleName = args[2];

        if (!AnsiConsole.Confirm($"Delete VBA module '[cyan]{moduleName}[/]'?"))
        {
            AnsiConsole.MarkupLine("[yellow]Cancelled[/]");
            return 1;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var deleteResult = await _coreCommands.DeleteAsync(batch, moduleName);
            await batch.SaveAsync();
            return deleteResult;
        });
        var result = task.GetAwaiter().GetResult();

        // Handle VBA trust guidance
        if (result is VbaTrustRequiredResult trustError)
        {
            DisplayVbaTrustGuidance(trustError);
            return 1;
        }

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Deleted VBA module '[cyan]{moduleName}[/]'");
        return 0;
    }
}
