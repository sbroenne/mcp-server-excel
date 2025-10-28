using Spectre.Console;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using System.Text;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Range operation commands - wraps Core with CLI CSV conversion
/// CLI converts CSV ↔ 2D arrays for user convenience (Core uses List&lt;List&lt;object?&gt;&gt;)
/// </summary>
public class RangeCommands
{
    private readonly Core.Commands.Range.RangeCommands _coreCommands = new();

    // === VALUE OPERATIONS ===

    public int GetValues(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-values <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-values data.xlsx Sheet1 A1:D10[/]");
            AnsiConsole.MarkupLine("[dim]Named range: range-get-values data.xlsx \"\" SalesData[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetValuesAsync(batch, sheetName, rangeAddress);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            // Output as CSV - transform rows to CSV format
            foreach (var csvRow in result.Values.Select(row => string.Join(",", row.Select(v => FormatCsvValue(v)))))
            {
                Console.WriteLine(csvRow);
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetValues(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-set-values <file.xlsx> <sheet-name> <range-address> <csv-file>");
            AnsiConsole.MarkupLine("[dim]Example: range-set-values data.xlsx Sheet1 A1:D10 input.csv[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var csvFile = args[4];

        if (!File.Exists(csvFile))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] CSV file not found: {csvFile}");
            return 1;
        }

        // Parse CSV to 2D array
        var csvData = File.ReadAllText(csvFile);
        var values = ParseCsvTo2DArray(csvData);

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var setResult = await _coreCommands.SetValuesAsync(batch, sheetName, rangeAddress, values);
            await batch.SaveAsync();
            return setResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set values in {rangeAddress.EscapeMarkup()}");

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

    // === FORMULA OPERATIONS ===

    public int GetFormulas(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-formulas <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-formulas data.xlsx Sheet1 C1:C10[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetFormulasAsync(batch, sheetName, rangeAddress);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            // Output formulas as CSV (empty string if no formula) - transform rows to CSV format
            foreach (var csvRow in result.Formulas.Select(row => string.Join(",", row.Select(f => FormatCsvValue(f)))))
            {
                Console.WriteLine(csvRow);
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetFormulas(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-set-formulas <file.xlsx> <sheet-name> <range-address> <csv-file>");
            AnsiConsole.MarkupLine("[dim]Example: range-set-formulas data.xlsx Sheet1 C1:C10 formulas.csv[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var csvFile = args[4];

        if (!File.Exists(csvFile))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] CSV file not found: {csvFile}");
            return 1;
        }

        // Parse CSV to 2D string array
        var csvData = File.ReadAllText(csvFile);
        var formulas = ParseCsvTo2DStringArray(csvData);

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var setResult = await _coreCommands.SetFormulasAsync(batch, sheetName, rangeAddress, formulas);
            await batch.SaveAsync();
            return setResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set formulas in {rangeAddress.EscapeMarkup()}");

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

    // === CLEAR OPERATIONS ===

    public int ClearAll(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-clear-all <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-clear-all data.xlsx Sheet1 A1:D10[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var clearResult = await _coreCommands.ClearAllAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return clearResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Cleared all content in {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ClearContents(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-clear-contents <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-clear-contents data.xlsx Sheet1 A1:D10[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var clearResult = await _coreCommands.ClearContentsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return clearResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Cleared contents in {rangeAddress.EscapeMarkup()} (formatting preserved)");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ClearFormats(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-clear-formats <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-clear-formats data.xlsx Sheet1 A1:D10[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var clearResult = await _coreCommands.ClearFormatsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return clearResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Cleared formats in {rangeAddress.EscapeMarkup()} (values preserved)");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === CSV CONVERSION HELPERS ===

    private static List<List<object?>> ParseCsvTo2DArray(string csvData)
    {
        var lines = csvData.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

        // Transform lines to rows using Select
        return lines.Select(line =>
        {
            var values = line.Split(',');
            return values.Select(value =>
            {
                var trimmed = value.Trim();

                // Try to parse as number
                if (double.TryParse(trimmed, out var number))
                {
                    return (object?)number;
                }
                // Try to parse as boolean
                else if (bool.TryParse(trimmed, out var boolean))
                {
                    return (object?)boolean;
                }
                // Empty string → null
                else if (string.IsNullOrEmpty(trimmed))
                {
                    return null;
                }
                // Otherwise string
                else
                {
                    return (object?)trimmed;
                }
            }).ToList();
        }).ToList();
    }

    private static List<List<string>> ParseCsvTo2DStringArray(string csvData)
    {
        var result = new List<List<string>>();
        var lines = csvData.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

        foreach (var line in lines)
        {
            var row = new List<string>();
            var values = line.Split(',');

            foreach (var value in values)
            {
                row.Add(value.Trim());
            }

            result.Add(row);
        }

        return result;
    }

    private static string FormatCsvValue(object? value)
    {
        if (value == null) return "";

        var str = value.ToString() ?? "";

        // If value contains comma, quote it
        if (str.Contains(',') || str.Contains('"') || str.Contains('\n'))
        {
            // Escape quotes
            str = str.Replace("\"", "\"\"");
            return $"\"{str}\"";
        }

        return str;
    }
}
