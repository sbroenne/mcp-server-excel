using Spectre.Console;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// VBA script management commands - CLI presentation layer (formats Core results)
/// </summary>
public class ScriptCommands : IScriptCommands
{
    private readonly Core.Commands.IScriptCommands _coreCommands;

    public ScriptCommands()
    {
        _coreCommands = new Core.Commands.ScriptCommands();
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

        var result = _coreCommands.List(filePath);

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
            else if (result.ErrorMessage?.Contains("not trusted") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Solution:[/] Run: [cyan]ExcelCLI setup-vba-trust[/]");
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

        var result = _coreCommands.Export(filePath, moduleName, outputFile).Result;

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

        var result = await _coreCommands.Import(filePath, moduleName, vbaFile);

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

        var result = await _coreCommands.Update(filePath, moduleName, vbaFile);

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

        var result = _coreCommands.Run(filePath, procedureName, parameters);

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            
            if (result.ErrorMessage?.Contains("not trusted") == true)
            {
                AnsiConsole.MarkupLine("[yellow]Solution:[/] Run: [cyan]ExcelCLI setup-vba-trust[/]");
            }
            
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

        var result = _coreCommands.Delete(filePath, moduleName);

        if (!result.Success)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }

        AnsiConsole.MarkupLine($"[green]✓[/] Deleted VBA module '[cyan]{moduleName}[/]'");
        return 0;
    }
}
