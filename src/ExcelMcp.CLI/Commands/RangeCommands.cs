using Sbroenne.ExcelMcp.ComInterop.Session;
using Spectre.Console;

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

    // === NUMBER FORMATTING OPERATIONS ===

    public int GetNumberFormats(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-number-formats <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-number-formats data.xlsx Sheet1 A1:D10[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.GetNumberFormatsAsync(batch, sheetName, rangeAddress);
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            // Output as CSV - transform rows to CSV format
            foreach (var csvRow in result.Formats.Select(row => string.Join(",", row.Select(v => FormatCsvValue(v)))))
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

    public int SetNumberFormat(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-set-number-format <file.xlsx> <sheet-name> <range-address> <format-code>");
            AnsiConsole.MarkupLine("[dim]Example: range-set-number-format data.xlsx Sheet1 D2:D100 \"$#,##0.00\"[/]");
            AnsiConsole.MarkupLine("[dim]Percentage: range-set-number-format data.xlsx Sheet1 E2:E100 \"0.00%\"[/]");
            AnsiConsole.MarkupLine("[dim]Date: range-set-number-format data.xlsx Sheet1 A2:A100 \"m/d/yyyy\"[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var formatCode = args[4];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var setResult = await _coreCommands.SetNumberFormatAsync(batch, sheetName, rangeAddress, formatCode);
            await batch.SaveAsync();
            return setResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Applied number format to {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === VISUAL FORMATTING OPERATIONS ===

    public int FormatRange(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-format <file.xlsx> <sheet-name> <range-address> [options]");
            AnsiConsole.MarkupLine("[dim]Font options: --font-name NAME --font-size SIZE --bold --italic --underline --font-color #RRGGBB[/]");
            AnsiConsole.MarkupLine("[dim]Fill options: --fill-color #RRGGBB[/]");
            AnsiConsole.MarkupLine("[dim]Border options: --border-style Continuous|Dashed|Dotted --border-weight Thin|Medium|Thick --border-color #RRGGBB[/]");
            AnsiConsole.MarkupLine("[dim]Alignment options: --h-align Left|Center|Right --v-align Top|Center|Bottom --wrap-text --orientation DEGREES[/]");
            AnsiConsole.MarkupLine("[dim]Example: range-format data.xlsx Sheet1 A1:E1 --bold --font-size 12 --h-align Center[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        // Parse formatting options from remaining args
        string? fontName = null;
        double? fontSize = null;
        bool? bold = null;
        bool? italic = null;
        bool? underline = null;
        string? fontColor = null;
        string? fillColor = null;
        string? borderStyle = null;
        string? borderColor = null;
        string? borderWeight = null;
        string? horizontalAlignment = null;
        string? verticalAlignment = null;
        bool? wrapText = null;
        int? orientation = null;

        for (int i = 4; i < args.Length; i++)
        {
            switch (args[i].ToLower())
            {
                case "--bold":
                    bold = true;
                    break;
                case "--italic":
                    italic = true;
                    break;
                case "--underline":
                    underline = true;
                    break;
                case "--wrap-text":
                    wrapText = true;
                    break;
                case "--font-name":
                    if (i + 1 < args.Length)
                    {
                        fontName = args[i + 1];
                        i++;
                    }
                    break;
                case "--font-size":
                    if (i + 1 < args.Length && double.TryParse(args[i + 1], out var size))
                    {
                        fontSize = size;
                        i++;
                    }
                    break;
                case "--font-color":
                    if (i + 1 < args.Length)
                    {
                        fontColor = args[i + 1];
                        i++;
                    }
                    break;
                case "--fill-color":
                    if (i + 1 < args.Length)
                    {
                        fillColor = args[i + 1];
                        i++;
                    }
                    break;
                case "--border-style":
                    if (i + 1 < args.Length)
                    {
                        borderStyle = args[i + 1];
                        i++;
                    }
                    break;
                case "--border-color":
                    if (i + 1 < args.Length)
                    {
                        borderColor = args[i + 1];
                        i++;
                    }
                    break;
                case "--border-weight":
                    if (i + 1 < args.Length)
                    {
                        borderWeight = args[i + 1];
                        i++;
                    }
                    break;
                case "--h-align":
                    if (i + 1 < args.Length)
                    {
                        horizontalAlignment = args[i + 1];
                        i++;
                    }
                    break;
                case "--v-align":
                    if (i + 1 < args.Length)
                    {
                        verticalAlignment = args[i + 1];
                        i++;
                    }
                    break;
                case "--orientation":
                    if (i + 1 < args.Length && int.TryParse(args[i + 1], out var degrees))
                    {
                        orientation = degrees;
                        i++;
                    }
                    break;
            }
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var setResult = await _coreCommands.FormatRangeAsync(batch, sheetName, rangeAddress,
                fontName, fontSize, bold, italic, underline, fontColor, fillColor,
                borderStyle, borderColor, borderWeight,
                horizontalAlignment, verticalAlignment, wrapText, orientation);
            await batch.SaveAsync();
            return setResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Applied formatting to {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === DATA VALIDATION OPERATIONS ===

    public int ValidateRange(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-validate <file.xlsx> <sheet-name> <range-address> <type> <formula1> [formula2] [options]");
            AnsiConsole.MarkupLine("[dim]Example (dropdown): range-validate data.xlsx Sheet1 F2:F100 List \"Active,Inactive,Pending\"[/]");
            AnsiConsole.MarkupLine("[dim]Example (number range): range-validate data.xlsx Sheet1 E2:E100 WholeNumber \"1\" \"999\" --operator Between[/]");
            AnsiConsole.MarkupLine("[dim]Types: List, WholeNumber, Decimal, Date, Time, TextLength, Custom[/]");
            AnsiConsole.MarkupLine("[dim]Operators: Between, NotBetween, Equal, NotEqual, Greater, Less, GreaterOrEqual, LessOrEqual[/]");
            AnsiConsole.MarkupLine("[dim]Options: --show-input --input-title TITLE --input-message MSG --error-title TITLE --error-message MSG[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var validationType = args[4];
        var formula1 = args[5];
        var formula2 = args.Length > 6 && !args[6].StartsWith("--") ? args[6] : null;

        // Parse optional parameters
        string? validationOperator = null;
        bool? showInputMessage = null;
        string? inputTitle = null;
        string? inputMessage = null;
        bool? showErrorAlert = null;
        string? errorStyle = null;
        string? errorTitle = null;
        string? errorMessage = null;
        bool? ignoreBlank = null;
        bool? showDropdown = null;

        int startIndex = formula2 != null ? 7 : 6;
        for (int i = startIndex; i < args.Length; i++)
        {
            switch (args[i].ToLower())
            {
                case "--operator":
                    if (i + 1 < args.Length)
                    {
                        validationOperator = args[i + 1];
                        i++;
                    }
                    break;
                case "--show-input":
                    showInputMessage = true;
                    break;
                case "--input-title":
                    if (i + 1 < args.Length)
                    {
                        inputTitle = args[i + 1];
                        i++;
                    }
                    break;
                case "--input-message":
                    if (i + 1 < args.Length)
                    {
                        inputMessage = args[i + 1];
                        i++;
                    }
                    break;
                case "--error-title":
                    if (i + 1 < args.Length)
                    {
                        errorTitle = args[i + 1];
                        i++;
                    }
                    break;
                case "--error-message":
                    if (i + 1 < args.Length)
                    {
                        errorMessage = args[i + 1];
                        i++;
                    }
                    break;
                case "--error-style":
                    if (i + 1 < args.Length)
                    {
                        errorStyle = args[i + 1];
                        i++;
                    }
                    break;
                case "--ignore-blank":
                    ignoreBlank = true;
                    break;
                case "--show-dropdown":
                    showDropdown = true;
                    break;
            }
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var setResult = await _coreCommands.ValidateRangeAsync(batch, sheetName, rangeAddress,
                validationType, validationOperator, formula1, formula2,
                showInputMessage, inputTitle, inputMessage,
                showErrorAlert, errorStyle, errorTitle, errorMessage,
                ignoreBlank, showDropdown);
            await batch.SaveAsync();
            return setResult;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Applied {validationType} validation to {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === HELPER METHODS ===

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
