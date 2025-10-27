using Spectre.Console;
using Sbroenne.ExcelMcp.Core.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Individual cell operation commands - wraps Core with CLI formatting
/// </summary>
public class CellCommands : ICellCommands
{
    private readonly Core.Commands.CellCommands _coreCommands = new();

    public int GetValue(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-get-value <file.xlsx> <sheet-name> <cell-address>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        CellValueResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                return await _coreCommands.GetValueAsync(batch, sheetName, cellAddress);
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
            string displayValue = result.Value?.ToString() ?? "[null]";
            AnsiConsole.MarkupLine($"[cyan]{result.CellAddress}:[/] {displayValue.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetValue(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-value <file.xlsx> <sheet-name> <cell-address> <value>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var value = args[4];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var setResult = await _coreCommands.SetValueAsync(batch, sheetName, cellAddress, value);
            await batch.SaveAsync();
            return setResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set {sheetName}!{cellAddress} = '{value.EscapeMarkup()}'");

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

    public int GetFormula(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-get-formula <file.xlsx> <sheet-name> <cell-address>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetFormulaAsync(batch, sheetName, cellAddress);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            string displayValue = result.Value?.ToString() ?? "[null]";

            if (string.IsNullOrEmpty(result.Formula))
            {
                AnsiConsole.MarkupLine($"[cyan]{result.CellAddress}:[/] [yellow](no formula)[/] Value: {displayValue.EscapeMarkup()}");
            }
            else
            {
                AnsiConsole.MarkupLine($"[cyan]{result.CellAddress}:[/] {result.Formula.EscapeMarkup()}");
                AnsiConsole.MarkupLine($"[dim]Result: {displayValue.EscapeMarkup()}[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetFormula(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cell-set-formula <file.xlsx> <sheet-name> <cell-address> <formula>");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var formula = args[4];

        OperationResult result;
        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var setResult = await _coreCommands.SetFormulaAsync(batch, sheetName, cellAddress, formula);
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
            // Need to get the result value by calling GetValue (separate batch)
            var valueTask = Task.Run(async () =>
            {
                await using var valueBatch = await ExcelSession.BeginBatchAsync(filePath);
                return await _coreCommands.GetValueAsync(valueBatch, sheetName, cellAddress);
            });
            var valueResult = valueTask.GetAwaiter().GetResult();
            string displayResult = valueResult.Value?.ToString() ?? "[null]";

            AnsiConsole.MarkupLine($"[green]✓[/] Set {sheetName}!{cellAddress} = {formula.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[dim]Result: {displayResult.EscapeMarkup()}[/]");

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
