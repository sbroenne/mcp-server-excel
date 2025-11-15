using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
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
            switch (args[i].ToLowerInvariant())
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
        var formula2 = args.Length > 6 && !args[6].StartsWith("--", StringComparison.Ordinal) ? args[6] : null;

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
            switch (args[i].ToLowerInvariant())
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

    // === COPY OPERATIONS ===

    public int CopyValues(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-copy-values <file.xlsx> <source-sheet> <source-range> <target-sheet> <target-range>");
            AnsiConsole.MarkupLine("[dim]Example: range-copy-values data.xlsx Sales A1:D100 Summary E5[/]");
            return 1;
        }

        var filePath = args[1];
        var sourceSheet = args[2];
        var sourceRange = args[3];
        var targetSheet = args[4];
        var targetRange = args[5];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.CopyValuesAsync(batch, sourceSheet, sourceRange, targetSheet, targetRange);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Copied values from {sourceRange.EscapeMarkup()} to {targetRange.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int CopyFormulas(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-copy-formulas <file.xlsx> <source-sheet> <source-range> <target-sheet> <target-range>");
            AnsiConsole.MarkupLine("[dim]Example: range-copy-formulas data.xlsx Template C1:E10 Sheet2 A1[/]");
            return 1;
        }

        var filePath = args[1];
        var sourceSheet = args[2];
        var sourceRange = args[3];
        var targetSheet = args[4];
        var targetRange = args[5];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.CopyFormulasAsync(batch, sourceSheet, sourceRange, targetSheet, targetRange);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Copied formulas from {sourceRange.EscapeMarkup()} to {targetRange.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === INSERT/DELETE OPERATIONS ===

    public int InsertCells(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-insert-cells <file.xlsx> <sheet-name> <range-address> <shift-direction>");
            AnsiConsole.MarkupLine("[dim]Example: range-insert-cells data.xlsx Sheet1 B5:C10 Down[/]");
            AnsiConsole.MarkupLine("[dim]Shift directions: Down, Right[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var shiftStr = args[4];

        if (!Enum.TryParse<InsertShiftDirection>(shiftStr, true, out var shift))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Invalid shift direction '{shiftStr.EscapeMarkup()}'. Use: Down, Right");
            return 1;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.InsertCellsAsync(batch, sheetName, rangeAddress, shift);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Inserted cells at {rangeAddress.EscapeMarkup()}, shifted {shiftStr}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int DeleteCells(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-delete-cells <file.xlsx> <sheet-name> <range-address> <shift-direction>");
            AnsiConsole.MarkupLine("[dim]Example: range-delete-cells data.xlsx Sheet1 B5:C10 Up[/]");
            AnsiConsole.MarkupLine("[dim]Shift directions: Up, Left[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var shiftStr = args[4];

        if (!Enum.TryParse<DeleteShiftDirection>(shiftStr, true, out var shift))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Invalid shift direction '{shiftStr.EscapeMarkup()}'. Use: Up, Left");
            return 1;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.DeleteCellsAsync(batch, sheetName, rangeAddress, shift);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted cells at {rangeAddress.EscapeMarkup()}, shifted {shiftStr}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int InsertRows(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-insert-rows <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-insert-rows data.xlsx Sheet1 5:7 (inserts 3 rows above row 5)[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.InsertRowsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Inserted rows above {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int DeleteRows(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-delete-rows <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-delete-rows data.xlsx Sheet1 5:7 (deletes rows 5-7)[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.DeleteRowsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted rows {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int InsertColumns(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-insert-columns <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-insert-columns data.xlsx Sheet1 C:E (inserts 3 columns before column C)[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.InsertColumnsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Inserted columns before {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int DeleteColumns(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-delete-columns <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-delete-columns data.xlsx Sheet1 C:E (deletes columns C-E)[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.DeleteColumnsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Deleted columns {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === FIND/REPLACE/SORT OPERATIONS ===

    public int Find(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-find <file.xlsx> <sheet-name> <range-address> <search-value> [--match-case] [--match-entire] [--search-formulas] [--search-values]");
            AnsiConsole.MarkupLine("[dim]Example: range-find data.xlsx Sheet1 A1:D100 \"error\" --match-case[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var searchValue = args[4];

        var options = new FindOptions();
        for (int i = 5; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--match-case":
                    options.MatchCase = true;
                    break;
                case "--match-entire":
                    options.MatchEntireCell = true;
                    break;
                case "--search-formulas":
                    options.SearchFormulas = true;
                    break;
                case "--search-values":
                    options.SearchValues = true;
                    break;
            }
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.FindAsync(batch, sheetName, rangeAddress, searchValue, options);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.MatchingCells.Count > 0)
            {
                AnsiConsole.MarkupLine($"[green]✓[/] Found {result.MatchingCells.Count} matches:");
                foreach (var cell in result.MatchingCells)
                {
                    AnsiConsole.MarkupLine($"  {cell.Address}: {cell.Value?.ToString()?.EscapeMarkup()}");
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No matches found[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Replace(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-replace <file.xlsx> <sheet-name> <range-address> <find-value> <replace-value> [--match-case] [--match-entire] [--replace-first]");
            AnsiConsole.MarkupLine("[dim]Example: range-replace data.xlsx Sheet1 A1:D100 \"old\" \"new\" --match-case[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var findValue = args[4];
        var replaceValue = args[5];

        var options = new ReplaceOptions { ReplaceAll = true };
        for (int i = 6; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--match-case":
                    options.MatchCase = true;
                    break;
                case "--match-entire":
                    options.MatchEntireCell = true;
                    break;
                case "--replace-first":
                    options.ReplaceAll = false;
                    break;
            }
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.ReplaceAsync(batch, sheetName, rangeAddress, findValue, replaceValue, options);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Replaced '{findValue.EscapeMarkup()}' with '{replaceValue.EscapeMarkup()}' in {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int Sort(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-sort <file.xlsx> <sheet-name> <range-address> <sort-spec> [--no-headers]");
            AnsiConsole.MarkupLine("[dim]Example: range-sort data.xlsx Sheet1 A1:D100 \"1:asc,3:desc\" (sort by col 1 ascending, then col 3 descending)[/]");
            AnsiConsole.MarkupLine("[dim]Sort spec format: \"colIndex:direction[,colIndex:direction...]\" where direction is asc or desc[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var sortSpec = args[4];
        var hasHeaders = true;

        if (args.Length > 5 && args[5] == "--no-headers")
        {
            hasHeaders = false;
        }

        // Parse sort specification
        var sortColumns = new List<SortColumn>();
        var specs = sortSpec.Split(',');
        foreach (var spec in specs)
        {
            var parts = spec.Split(':');
            if (parts.Length == 2 && int.TryParse(parts[0], out var colIndex))
            {
                var ascending = parts[1].Trim().Equals("asc", StringComparison.OrdinalIgnoreCase);
                sortColumns.Add(new SortColumn { ColumnIndex = colIndex, Ascending = ascending });
            }
        }

        if (sortColumns.Count == 0)
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Invalid sort specification: {sortSpec.EscapeMarkup()}");
            return 1;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SortAsync(batch, sheetName, rangeAddress, sortColumns, hasHeaders);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Sorted {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === DISCOVERY OPERATIONS ===

    public int GetUsedRange(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-used <file.xlsx> <sheet-name>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-used data.xlsx Sheet1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.GetUsedRangeAsync(batch, sheetName);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Used range: {result.RangeAddress.EscapeMarkup()} ({result.RowCount} rows × {result.ColumnCount} columns)");

            // Display values in table format
            var table = new Table();
            table.Border(TableBorder.Rounded);

            // Add columns
            for (int i = 0; i < result.ColumnCount; i++)
            {
                table.AddColumn($"Col {i + 1}");
            }

            // Add rows (limit to first 50 for readability)
            var displayRows = Math.Min(result.Values.Count, 50);
            for (int i = 0; i < displayRows; i++)
            {
                var row = result.Values[i];
                var cellValues = row.Select(FormatCsvValue).ToArray();
                table.AddRow(cellValues);
            }

            if (result.Values.Count > displayRows)
            {
                AnsiConsole.MarkupLine($"[dim]... and {result.Values.Count - displayRows} more rows[/]");
            }

            AnsiConsole.Write(table);
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetCurrentRegion(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-current-region <file.xlsx> <sheet-name> <cell-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-current-region data.xlsx Sheet1 B5 (finds contiguous data block around B5)[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.GetCurrentRegionAsync(batch, sheetName, cellAddress);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Current region around {cellAddress.EscapeMarkup()}: {result.RangeAddress.EscapeMarkup()} ({result.RowCount} rows × {result.ColumnCount} columns)");

            // Display values in table format
            var table = new Table();
            table.Border(TableBorder.Rounded);

            // Add columns
            for (int i = 0; i < result.ColumnCount; i++)
            {
                table.AddColumn($"Col {i + 1}");
            }

            // Add rows (limit to first 20 for readability)
            var displayRows = Math.Min(result.Values.Count, 20);
            for (int i = 0; i < displayRows; i++)
            {
                var row = result.Values[i];
                var cellValues = row.Select(FormatCsvValue).ToArray();
                table.AddRow(cellValues);
            }

            if (result.Values.Count > displayRows)
            {
                AnsiConsole.MarkupLine($"[dim]... and {result.Values.Count - displayRows} more rows[/]");
            }

            AnsiConsole.Write(table);
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetInfo(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-info <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-info data.xlsx Sheet1 A1:D100[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.GetInfoAsync(batch, sheetName, rangeAddress);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Range information:");
            AnsiConsole.MarkupLine($"  Address: {result.Address?.EscapeMarkup()}");
            AnsiConsole.MarkupLine($"  Dimensions: {result.RowCount} rows × {result.ColumnCount} columns");
            if (!string.IsNullOrEmpty(result.NumberFormat))
            {
                AnsiConsole.MarkupLine($"  Number format: {result.NumberFormat.EscapeMarkup()}");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === HYPERLINK OPERATIONS ===

    public int AddHyperlink(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-add-hyperlink <file.xlsx> <sheet-name> <cell-address> <url> [display-text] [tooltip]");
            AnsiConsole.MarkupLine("[dim]Example: range-add-hyperlink data.xlsx Sheet1 A1 \"https://example.com\" \"Click here\"[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];
        var url = args[4];
        var displayText = args.Length > 5 ? args[5] : null;
        var tooltip = args.Length > 6 ? args[6] : null;

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.AddHyperlinkAsync(batch, sheetName, cellAddress, url, displayText, tooltip);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Added hyperlink to {cellAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int RemoveHyperlink(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-remove-hyperlink <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-remove-hyperlink data.xlsx Sheet1 A1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.RemoveHyperlinkAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Removed hyperlink from {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ListHyperlinks(string[] args)
    {
        if (args.Length < 3)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-list-hyperlinks <file.xlsx> <sheet-name>");
            AnsiConsole.MarkupLine("[dim]Example: range-list-hyperlinks data.xlsx Sheet1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.ListHyperlinksAsync(batch, sheetName);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.Hyperlinks.Count > 0)
            {
                var table = new Table();
                table.Border(TableBorder.Rounded);
                table.AddColumn("Cell");
                table.AddColumn("URL");
                table.AddColumn("Display Text");

                foreach (var link in result.Hyperlinks)
                {
                    table.AddRow(
                        link.CellAddress.EscapeMarkup(),
                        link.Address?.EscapeMarkup() ?? "",
                        link.DisplayText?.EscapeMarkup() ?? ""
                    );
                }

                AnsiConsole.Write(table);
                AnsiConsole.MarkupLine($"[green]Found {result.Hyperlinks.Count} hyperlinks[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No hyperlinks found[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetHyperlink(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-hyperlink <file.xlsx> <sheet-name> <cell-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-hyperlink data.xlsx Sheet1 A1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var cellAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.GetHyperlinkAsync(batch, sheetName, cellAddress);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            if (result.Hyperlinks.Count > 0)
            {
                var link = result.Hyperlinks[0];
                AnsiConsole.MarkupLine($"[green]✓[/] Hyperlink in {cellAddress.EscapeMarkup()}:");
                AnsiConsole.MarkupLine($"  URL: {link.Address?.EscapeMarkup()}");
                if (!string.IsNullOrEmpty(link.DisplayText))
                {
                    AnsiConsole.MarkupLine($"  Display text: {link.DisplayText.EscapeMarkup()}");
                }
                if (!string.IsNullOrEmpty(link.ScreenTip))
                {
                    AnsiConsole.MarkupLine($"  Tooltip: {link.ScreenTip.EscapeMarkup()}");
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No hyperlink found[/]");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === STYLE OPERATIONS ===

    public int GetStyle(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-style <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-style data.xlsx Sheet1 A1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.GetStyleAsync(batch, sheetName, rangeAddress);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Style: {result.StyleName?.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int SetStyle(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-set-style <file.xlsx> <sheet-name> <range-address> <style-name>");
            AnsiConsole.MarkupLine("[dim]Example: range-set-style data.xlsx Sheet1 A1:D10 \"Heading 1\"[/]");
            AnsiConsole.MarkupLine("[dim]Common styles: Normal, Heading 1-4, Title, Total, Currency, Percent, Good, Bad, Neutral[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var styleName = args[4];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SetStyleAsync(batch, sheetName, rangeAddress, styleName);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Applied style '{styleName.EscapeMarkup()}' to {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === VALIDATION/AUTOFIT OPERATIONS ===

    public int GetValidation(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-validation <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-validation data.xlsx Sheet1 A1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.GetValidationAsync(batch, sheetName, rangeAddress);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Validation settings:");
            AnsiConsole.MarkupLine($"  Type: {result.ValidationType?.EscapeMarkup()}");
            if (!string.IsNullOrEmpty(result.ValidationOperator))
            {
                AnsiConsole.MarkupLine($"  Operator: {result.ValidationOperator.EscapeMarkup()}");
            }
            if (!string.IsNullOrEmpty(result.Formula1))
            {
                AnsiConsole.MarkupLine($"  Formula1: {result.Formula1.EscapeMarkup()}");
            }
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int RemoveValidation(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-remove-validation <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-remove-validation data.xlsx Sheet1 A1:A100[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.RemoveValidationAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Removed validation from {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int AutoFitColumns(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-autofit-columns <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-autofit-columns data.xlsx Sheet1 A:D[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.AutoFitColumnsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Auto-fitted columns in {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int AutoFitRows(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-autofit-rows <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-autofit-rows data.xlsx Sheet1 1:100[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.AutoFitRowsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Auto-fitted rows in {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === MERGE OPERATIONS ===

    public int MergeCells(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-merge-cells <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-merge-cells data.xlsx Sheet1 A1:C1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.MergeCellsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Merged cells {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int UnmergeCells(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-unmerge-cells <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-unmerge-cells data.xlsx Sheet1 A1:C1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.UnmergeCellsAsync(batch, sheetName, rangeAddress);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Unmerged cells {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetMergeInfo(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-merge-info <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-merge-info data.xlsx Sheet1 A1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.GetMergeInfoAsync(batch, sheetName, rangeAddress);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Merge status: {(result.IsMerged ? "Merged" : "Not merged")}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    // === ADVANCED OPERATIONS ===

    public int SetCellLock(string[] args)
    {
        if (args.Length < 5)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-set-cell-lock <file.xlsx> <sheet-name> <range-address> <locked>");
            AnsiConsole.MarkupLine("[dim]Example: range-set-cell-lock data.xlsx Sheet1 A1:D10 true[/]");
            AnsiConsole.MarkupLine("[dim]Note: Cell lock only takes effect when worksheet protection is enabled[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        if (!bool.TryParse(args[4], out var locked))
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] Invalid locked value '{args[4]}'. Use: true or false");
            return 1;
        }

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.SetCellLockAsync(batch, sheetName, rangeAddress, locked);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Set cell lock to {locked} for {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int GetCellLock(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-get-cell-lock <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-get-cell-lock data.xlsx Sheet1 A1[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.GetCellLockAsync(batch, sheetName, rangeAddress);
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Cell lock status: {(result.IsLocked ? "Locked" : "Unlocked")}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int AddConditionalFormatting(string[] args)
    {
        if (args.Length < 6)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-add-conditional-formatting <file.xlsx> <sheet-name> <range-address> <rule-type> <formula1> [formula2] [format-style]");
            AnsiConsole.MarkupLine("[dim]Example: range-add-conditional-formatting data.xlsx Sheet1 A1:A10 cellValue \">100\" \"\" highlight[/]");
            AnsiConsole.MarkupLine("[dim]Rule types: cellValue, expression, colorScale, dataBar, iconSet, top10, uniqueValues, duplicateValues[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];
        var ruleType = args[4];
        var formula1 = args[5];
        var formula2 = args.Length > 6 ? args[6] : null;
        var formatStyle = args.Length > 7 ? args[7] : null;

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.AddConditionalFormattingAsync(batch, sheetName, rangeAddress, ruleType, formula1, formula2, formatStyle);
            await batch.SaveAsync();
            return result;
        });
        var result = task.GetAwaiter().GetResult();

        if (result.Success)
        {
            AnsiConsole.MarkupLine($"[green]✓[/] Added conditional formatting to {rangeAddress.EscapeMarkup()}");
            return 0;
        }
        else
        {
            AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage?.EscapeMarkup()}");
            return 1;
        }
    }

    public int ClearConditionalFormatting(string[] args)
    {
        if (args.Length < 4)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] range-clear-conditional-formatting <file.xlsx> <sheet-name> <range-address>");
            AnsiConsole.MarkupLine("[dim]Example: range-clear-conditional-formatting data.xlsx Sheet1 A1:A10[/]");
            return 1;
        }

        var filePath = args[1];
        var sheetName = args[2];
        var rangeAddress = args[3];

        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var result = await _coreCommands.ClearConditionalFormattingAsync(batch, sheetName, rangeAddress);
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

    // === HELPER METHODS ===



    private static readonly string[] LineSeparators = ["\r\n", "\r", "\n"];

    private static List<List<object?>> ParseCsvTo2DArray(string csvData)
    {
        var lines = csvData.Split(LineSeparators, StringSplitOptions.RemoveEmptyEntries);

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
        var lines = csvData.Split(LineSeparators, StringSplitOptions.RemoveEmptyEntries);

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
