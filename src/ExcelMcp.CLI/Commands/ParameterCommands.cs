using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Named range/parameter management commands - wraps Core with CLI formatting
/// </summary>
public class NamedRangeCommands : INamedRangeCommands
{
    private readonly Core.Commands.NamedRangeCommands _coreCommands = new();

    public int List(string[] args)
    {
        if (args.Length < 2)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] param-list <file.xlsx>");
            return 1;
        }

        var filePath = args[1];
        AnsiConsole.MarkupLine($"[bold]Named Ranges/Parameters in:[/] {Path.GetFileName(filePath)}\n");

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.ListAsync(batch);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.NamedRanges.Count > 0)
            {
                var table = new Table();
                table.AddColumn("[bold]Parameter Name[/]");
                table.AddColumn("[bold]Refers To[/]");
                table.AddColumn("[bold]Value[/]");

                foreach (var param in result.NamedRanges.OrderBy(p => p.Name))
                {
                    string refersTo = param.RefersTo.Length > 40 ? param.RefersTo[..37] + "..." : param.RefersTo;
                    string value = param.Value?.ToString() ?? "[null]";
                    table.AddRow(param.Name.EscapeMarkup(), refersTo.EscapeMarkup(), value.EscapeMarkup());
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"\n[dim]Found {result.NamedRanges.Count} parameter(s)[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No named ranges found in this workbook[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Set(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] param-set <file.xlsx> <param-name> <value>");
            return 1;
        }

        var filePath = args[1];
        var paramName = args[2];
        var value = args[3];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var setResult = await _coreCommands.SetAsync(batch, paramName, value);
                await batch.SaveAsync();
                return setResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set parameter '{paramName.EscapeMarkup()}' = '{value.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Get(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] param-get <file.xlsx> <param-name>");
            return 1;
        }

        var filePath = args[1];
        var paramName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetAsync(batch, paramName);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            string value = result.Value?.ToString() ?? "[null]";
            AnsiConsole.MarkupLine($"[cyan]{paramName}:[/] {value.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[dim]Refers to: {result.RefersTo.EscapeMarkup()}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Update(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] param-update <file.xlsx> <param-name> <new-reference>");
            AnsiConsole.MarkupLine("[yellow]Example:[/] param-update data.xlsx MyParam Config!B5");
            AnsiConsole.MarkupLine("\n[dim]Note: This updates the cell reference, not the value.[/]");
            AnsiConsole.MarkupLine("[dim]Use 'param-set' to change the value.[/]");
            return 1;
        }

        var filePath = args[1];
        var paramName = args[2];
        var reference = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var updateResult = await _coreCommands.UpdateAsync(batch, paramName, reference);
            await batch.SaveAsync();
            return updateResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Updated parameter '{paramName.EscapeMarkup()}' reference to {reference.EscapeMarkup()}");
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
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] param-create <file.xlsx> <param-name> <reference>");
            AnsiConsole.MarkupLine("[yellow]Example:[/] param-create data.xlsx MyParam Sheet1!A1");
            return 1;
        }

        var filePath = args[1];
        var paramName = args[2];
        var reference = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var createResult = await _coreCommands.CreateAsync(batch, paramName, reference);
            await batch.SaveAsync();
            return createResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Created parameter '{paramName.EscapeMarkup()}' -> {reference.EscapeMarkup()}");
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
            AnsiConsole.MarkupLine("[red]Usage:[/] param-delete <file.xlsx> <param-name>");
            return 1;
        }

        var filePath = args[1];
        var paramName = args[2];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var deleteResult = await _coreCommands.DeleteAsync(batch, paramName);
                await batch.SaveAsync();
                return deleteResult;
            });
            result = task.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted parameter '{paramName.EscapeMarkup()}'");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}
