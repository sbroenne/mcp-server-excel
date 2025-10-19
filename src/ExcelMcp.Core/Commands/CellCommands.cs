using Spectre.Console;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Individual cell operation commands implementation
/// </summary>
public class CellCommands : ICellCommands
{
    /// <inheritdoc />
    public int GetValue(string[] args)
    {
        if (!ValidateArgs(args, 4, "cell-get-value <file.xlsx> <sheet-name> <cell-address>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var sheetName = args[2];
        var cellAddress = args[3];

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName}' not found");
                    return 1;
                }

                dynamic cell = sheet.Range[cellAddress];
                object value = cell.Value2;
                string displayValue = value?.ToString() ?? "[null]";

                AnsiConsole.MarkupLine($"[cyan]{sheetName}!{cellAddress}:[/] {displayValue.EscapeMarkup()}");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    /// <inheritdoc />
    public int SetValue(string[] args)
    {
        if (!ValidateArgs(args, 5, "cell-set-value <file.xlsx> <sheet-name> <cell-address> <value>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var sheetName = args[2];
        var cellAddress = args[3];
        var value = args[4];

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName}' not found");
                    return 1;
                }

                dynamic cell = sheet.Range[cellAddress];
                
                // Try to parse as number, otherwise set as text
                if (double.TryParse(value, out double numValue))
                {
                    cell.Value2 = numValue;
                }
                else if (bool.TryParse(value, out bool boolValue))
                {
                    cell.Value2 = boolValue;
                }
                else
                {
                    cell.Value2 = value;
                }

                workbook.Save();
                AnsiConsole.MarkupLine($"[green]✓[/] Set {sheetName}!{cellAddress} = '{value.EscapeMarkup()}'");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    /// <inheritdoc />
    public int GetFormula(string[] args)
    {
        if (!ValidateArgs(args, 4, "cell-get-formula <file.xlsx> <sheet-name> <cell-address>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var sheetName = args[2];
        var cellAddress = args[3];

        return WithExcel(args[1], false, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName}' not found");
                    return 1;
                }

                dynamic cell = sheet.Range[cellAddress];
                string formula = cell.Formula ?? "";
                object value = cell.Value2;
                string displayValue = value?.ToString() ?? "[null]";

                if (string.IsNullOrEmpty(formula))
                {
                    AnsiConsole.MarkupLine($"[cyan]{sheetName}!{cellAddress}:[/] [yellow](no formula)[/] Value: {displayValue.EscapeMarkup()}");
                }
                else
                {
                    AnsiConsole.MarkupLine($"[cyan]{sheetName}!{cellAddress}:[/] {formula.EscapeMarkup()}");
                    AnsiConsole.MarkupLine($"[dim]Result: {displayValue.EscapeMarkup()}[/]");
                }

                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }

    /// <inheritdoc />
    public int SetFormula(string[] args)
    {
        if (!ValidateArgs(args, 5, "cell-set-formula <file.xlsx> <sheet-name> <cell-address> <formula>")) return 1;
        if (!File.Exists(args[1]))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {args[1]}");
            return 1;
        }

        var sheetName = args[2];
        var cellAddress = args[3];
        var formula = args[4];

        // Ensure formula starts with =
        if (!formula.StartsWith("="))
        {
            formula = "=" + formula;
        }

        return WithExcel(args[1], true, (excel, workbook) =>
        {
            try
            {
                dynamic? sheet = FindSheet(workbook, sheetName);
                if (sheet == null)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] Sheet '{sheetName}' not found");
                    return 1;
                }

                dynamic cell = sheet.Range[cellAddress];
                cell.Formula = formula;

                workbook.Save();
                
                // Get the calculated result
                object result = cell.Value2;
                string displayResult = result?.ToString() ?? "[null]";
                
                AnsiConsole.MarkupLine($"[green]✓[/] Set {sheetName}!{cellAddress} = {formula.EscapeMarkup()}");
                AnsiConsole.MarkupLine($"[dim]Result: {displayResult.EscapeMarkup()}[/]");
                return 0;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
                return 1;
            }
        });
    }
}
