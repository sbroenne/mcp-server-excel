using Sbroenne.ExcelMcp.ComInterop.Session;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Conditional formatting commands - wraps Core ConditionalFormattingCommands with CLI interface
/// </summary>
public class ConditionalFormatCommands
{
    private readonly Core.Commands.ConditionalFormattingCommands _coreCommands = new();

    public int AddRule(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cf-add-rule <file.xlsx> <sheet-name> <range-address> <rule-type> <operator-type> <formula1> [formula2] [interior-color] [interior-pattern] [font-color] [font-bold] [font-italic] [border-style] [border-color]");
            AnsiConsole.MarkupLine("[dim]Example: cf-add-rule data.xlsx Sheet1 A1:A10 cell-value greater 100 \"\" #FFFF00 solid[/]");
            AnsiConsole.MarkupLine("[dim]Rule types: cell-value, expression[/]");
            AnsiConsole.MarkupLine("[dim]Operators: equal, not-equal, greater, less, greater-equal, less-equal, between, not-between[/]");
            AnsiConsole.MarkupLine("[dim]Colors: #RRGGBB hex or color index[/]");
            AnsiConsole.MarkupLine("[dim]Patterns: solid, gray75, gray50, gray25, horizontal, vertical, etc.[/]");
            AnsiConsole.MarkupLine("[dim]Border styles: continuous, dash, dot, dash-dot, dash-dot-dot, double[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var ruleType = args[4];
        var operatorType = args.Length > 5 ? args[5] : null;
        var formula1 = args.Length > 6 ? args[6] : null;
        var formula2 = args.Length > 7 && !string.IsNullOrEmpty(args[7]) ? args[7] : null;
        var interiorColor = args.Length > 8 && !string.IsNullOrEmpty(args[8]) ? args[8] : null;
        var interiorPattern = args.Length > 9 && !string.IsNullOrEmpty(args[9]) ? args[9] : null;
        var fontColor = args.Length > 10 && !string.IsNullOrEmpty(args[10]) ? args[10] : null;
        var fontBold = args.Length > 11 ? bool.TryParse(args[11], out var bold) ? bold : (bool?)null : null;
        var fontItalic = args.Length > 12 ? bool.TryParse(args[12], out var italic) ? italic : (bool?)null : null;
        var borderStyle = args.Length > 13 && !string.IsNullOrEmpty(args[13]) ? args[13] : null;
        var borderColor = args.Length > 14 && !string.IsNullOrEmpty(args[14]) ? args[14] : null;

        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var result = await _coreCommands.AddRuleAsync(
                    batch, sheetName, rangeAddress, ruleType, operatorType, formula1, formula2,
                    interiorColor, interiorPattern, fontColor, fontBold, fontItalic, borderStyle, borderColor);
                await batch.SaveAsync();
                return result;
            });
            var result = task.GetAwaiter().GetResult();

            if (result.Success)
            {
                AnsiConsole.MarkupLine($"[green]✓[/] Added conditional formatting rule to {rangeAddress.EscapeMarkup()}");
                return 0;
            }
            else
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
                return 1;
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }

    public int ClearRules(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] cf-clear-rules <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: cf-clear-rules data.xlsx Sheet1 A1:A10[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        try
        {
            var task = Task.Run(async () =>
            {
                await using var batch = await ExcelSession.BeginBatchAsync(filePath);
                var result = await _coreCommands.ClearRulesAsync(batch, sheetName, rangeAddress);
                await batch.SaveAsync();
                return result;
            });
            var result = task.GetAwaiter().GetResult();

            if (result.Success)
            {
                AnsiConsole.MarkupLine($"[green]✓[/] Cleared conditional formatting from {rangeAddress.EscapeMarkup()}");
                return 0;
            }
            else
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
                return 1;
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }
}
