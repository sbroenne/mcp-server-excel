using Spectre.Console;

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

        var result = _coreCommands.GetValue(filePath, sheetName, cellAddress);
        
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

        var result = _coreCommands.SetValue(filePath, sheetName, cellAddress, value);
        
        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set {sheetName}!{cellAddress} = '{value.EscapeMarkup()}'");
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

        var result = _coreCommands.GetFormula(filePath, sheetName, cellAddress);
        
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

        var result = _coreCommands.SetFormula(filePath, sheetName, cellAddress, formula);
        
        if (result.Success)
        {
            // Need to get the result value by calling GetValue
            var valueResult = _coreCommands.GetValue(filePath, sheetName, cellAddress);
            string displayResult = valueResult.Value?.ToString() ?? "[null]";
            
            AnsiConsole.MarkupLine($"[green]✓[/] Set {sheetName}!{cellAddress} = {formula.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"[dim]Result: {displayResult.EscapeMarkup()}[/]");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }
}