using System.ComponentModel;
using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Service;
using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Sbroenne.ExcelMcp.Core.Models.Actions;
using Spectre.Console;
using Spectre.Console.Cli;

namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Range commands - thin wrapper that sends requests to service.
/// </summary>
internal sealed class RangeCommand : AsyncCommand<RangeCommand.Settings>
{
    public override async Task<int> ExecuteAsync(CommandContext context, Settings settings, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(settings.SessionId))
        {
            AnsiConsole.MarkupLine("[red]Session ID is required. Use --session <id>[/]");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(settings.Action))
        {
            AnsiConsole.MarkupLine("[red]Action is required.[/]");
            return 1;
        }

        var validActions = ActionValidator.GetValidActions<RangeAction>()
            .Concat(ActionValidator.GetValidActions<RangeEditAction>())
            .Concat(ActionValidator.GetValidActions<RangeFormatAction>())
            .Concat(ActionValidator.GetValidActions<RangeLinkAction>())
            .ToArray();

        if (!ActionValidator.TryNormalizeAction(settings.Action, validActions, out var action, out var errorMessage))
        {
            AnsiConsole.MarkupLine($"[red]{errorMessage}[/]");
            return 1;
        }
        var command = $"range.{action}";

        // Build args based on action
        var values = ResolveFileOrValue(settings.Values, settings.ValuesFile);
        var formulas = ResolveFileOrValue(settings.Formulas, settings.FormulasFile);
        var formats = ResolveFileOrValue(settings.Formats, settings.FormatsFile);
        object? args = action switch
        {
            // RangeAction
            "get-values" => new { sheetName = settings.SheetName, range = settings.Range },
            "set-values" => BuildSetValuesArgs(settings, values),
            "get-formulas" => new { sheetName = settings.SheetName, range = settings.Range },
            "set-formulas" => new { sheetName = settings.SheetName, range = settings.Range, formulas = ParseStringArray(formulas) },
            "get-number-formats" => new { sheetName = settings.SheetName, range = settings.Range },
            "set-number-format" => new { sheetName = settings.SheetName, range = settings.Range, formatCode = settings.FormatCode },
            "set-number-formats" => new { sheetName = settings.SheetName, range = settings.Range, formats = ParseStringArray(formats) },
            "clear-all" => new { sheetName = settings.SheetName, range = settings.Range },
            "clear-contents" => new { sheetName = settings.SheetName, range = settings.Range },
            "clear-formats" => new { sheetName = settings.SheetName, range = settings.Range },
            "copy" => new { sourceSheet = settings.SourceSheet ?? settings.SheetName, sourceRange = settings.SourceRange ?? settings.Range, targetSheet = settings.TargetSheet, targetRange = settings.TargetRange },
            "copy-values" => new { sourceSheet = settings.SourceSheet ?? settings.SheetName, sourceRange = settings.SourceRange ?? settings.Range, targetSheet = settings.TargetSheet, targetRange = settings.TargetRange },
            "copy-formulas" => new { sourceSheet = settings.SourceSheet ?? settings.SheetName, sourceRange = settings.SourceRange ?? settings.Range, targetSheet = settings.TargetSheet, targetRange = settings.TargetRange },
            "get-used-range" => new { sheetName = settings.SheetName },
            "get-current-region" => new { sheetName = settings.SheetName, cellAddress = settings.CellAddress ?? settings.Range },
            "get-info" => new { sheetName = settings.SheetName, range = settings.Range },

            // RangeEditAction
            "insert-cells" => new { sheetName = settings.SheetName, range = settings.Range, shiftDirection = settings.ShiftDirection },
            "delete-cells" => new { sheetName = settings.SheetName, range = settings.Range, shiftDirection = settings.ShiftDirection },
            "insert-rows" => new { sheetName = settings.SheetName, range = settings.Range },
            "delete-rows" => new { sheetName = settings.SheetName, range = settings.Range },
            "insert-columns" => new { sheetName = settings.SheetName, range = settings.Range },
            "delete-columns" => new { sheetName = settings.SheetName, range = settings.Range },
            "find" => new { sheetName = settings.SheetName, range = settings.Range, searchValue = settings.SearchValue, matchCase = settings.MatchCase, matchEntireCell = settings.MatchEntireCell, searchFormulas = settings.SearchFormulas },
            "replace" => new { sheetName = settings.SheetName, range = settings.Range, findValue = settings.SearchValue, replaceValue = settings.ReplaceValue, matchCase = settings.MatchCase, matchEntireCell = settings.MatchEntireCell, replaceAll = settings.ReplaceAll },
            "sort" => new { sheetName = settings.SheetName, range = settings.Range, sortColumnsJson = settings.SortColumnsJson, hasHeaders = settings.HasHeaders },

            // RangeFormatAction
            "get-style" => new { sheetName = settings.SheetName, range = settings.Range },
            "set-style" => new { sheetName = settings.SheetName, range = settings.Range, styleName = settings.StyleName },
            "format-range" => new { sheetName = settings.SheetName, range = settings.Range, fontName = settings.FontName, fontSize = settings.FontSize, bold = settings.Bold, italic = settings.Italic, underline = settings.Underline, fontColor = settings.FontColor, fillColor = settings.FillColor, borderStyle = settings.BorderStyle, borderColor = settings.BorderColor, borderWeight = settings.BorderWeight, horizontalAlignment = settings.HorizontalAlignment, verticalAlignment = settings.VerticalAlignment, wrapText = settings.WrapText, orientation = settings.Orientation },
            "validate-range" => new { sheetName = settings.SheetName, range = settings.Range, validationType = settings.ValidationType, validationOperator = settings.ValidationOperator, formula1 = settings.Formula1, formula2 = settings.Formula2, showInputMessage = settings.ShowInputMessage, inputTitle = settings.InputTitle, inputMessage = settings.InputMessage, showErrorAlert = settings.ShowErrorAlert, errorStyle = settings.ErrorStyle, errorTitle = settings.ErrorTitle, errorMessage = settings.ErrorMessage, ignoreBlank = settings.IgnoreBlank, showDropdown = settings.ShowDropdown },
            "get-validation" => new { sheetName = settings.SheetName, range = settings.Range },
            "remove-validation" => new { sheetName = settings.SheetName, range = settings.Range },
            "auto-fit-columns" => new { sheetName = settings.SheetName, range = settings.Range },
            "auto-fit-rows" => new { sheetName = settings.SheetName, range = settings.Range },
            "merge-cells" => new { sheetName = settings.SheetName, range = settings.Range },
            "unmerge-cells" => new { sheetName = settings.SheetName, range = settings.Range },
            "get-merge-info" => new { sheetName = settings.SheetName, range = settings.Range },

            // RangeLinkAction
            "add-hyperlink" => new { sheetName = settings.SheetName, cellAddress = settings.CellAddress ?? settings.Range, url = settings.Url, displayText = settings.DisplayText, tooltip = settings.Tooltip },
            "remove-hyperlink" => new { sheetName = settings.SheetName, cellAddress = settings.CellAddress ?? settings.Range },
            "list-hyperlinks" => new { sheetName = settings.SheetName, range = settings.Range },
            "get-hyperlink" => new { sheetName = settings.SheetName, cellAddress = settings.CellAddress ?? settings.Range },
            "set-cell-lock" => new { sheetName = settings.SheetName, range = settings.Range, locked = settings.Locked },
            "get-cell-lock" => new { sheetName = settings.SheetName, range = settings.Range },

            _ => new { sheetName = settings.SheetName, range = settings.Range }
        };

        using var client = new ServiceClient();
        var response = await client.SendAsync(new ServiceRequest
        {
            Command = command,
            SessionId = settings.SessionId,
            Args = args != null ? JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions) : null
        }, cancellationToken);

        if (response.Success)
        {
            if (!string.IsNullOrEmpty(response.Result))
            {
                Console.WriteLine(response.Result);
            }
            else
            {
                Console.WriteLine(JsonSerializer.Serialize(new { success = true }, ServiceProtocol.JsonOptions));
            }
            return 0;
        }
        else
        {
            Console.WriteLine(JsonSerializer.Serialize(new { success = false, error = response.ErrorMessage }, ServiceProtocol.JsonOptions));
            return 1;
        }
    }

    private static object? BuildSetValuesArgs(Settings settings, string? valuesJson)
    {
        // Parse values from JSON string or file
        List<List<object?>>? values = null;
        if (!string.IsNullOrEmpty(valuesJson))
        {
            try
            {
                values = JsonSerializer.Deserialize<List<List<object?>>>(valuesJson, ServiceProtocol.JsonOptions);
            }
            catch
            {
                // If not valid JSON array, treat as single value
                values = [[valuesJson]];
            }
        }

        return new
        {
            sheetName = settings.SheetName,
            range = settings.Range,
            values
        };
    }

    private static List<List<string>>? ParseStringArray(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return null;
        try
        {
            return JsonSerializer.Deserialize<List<List<string>>>(input, ServiceProtocol.JsonOptions);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Returns file contents if filePath is provided, otherwise returns the direct value.
    /// </summary>
    private static string? ResolveFileOrValue(string? directValue, string? filePath)
    {
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found: {filePath}");
            }
            return File.ReadAllText(filePath);
        }
        return directValue;
    }

    internal sealed class Settings : CommandSettings
    {
        [CommandArgument(0, "<ACTION>")]
        [Description("The action to perform")]
        public string Action { get; init; } = string.Empty;

        [CommandOption("-s|--session <SESSION>")]
        [Description("Session ID from 'session open' command")]
        public string SessionId { get; init; } = string.Empty;

        [CommandOption("--sheet <NAME>")]
        [Description("Target worksheet name")]
        public string? SheetName { get; init; }

        [CommandOption("--range <ADDRESS>")]
        [Description("Cell range address (e.g., A1:C10)")]
        public string? Range { get; init; }

        [CommandOption("--cell <ADDRESS>")]
        [Description("Single cell address")]
        public string? CellAddress { get; init; }

        // Value/Formula options
        [CommandOption("--values <JSON>")]
        [Description("Cell values as JSON 2D array")]
        public string? Values { get; init; }

        [CommandOption("--values-file <PATH>")]
        [Description("Read values JSON from file")]
        public string? ValuesFile { get; init; }

        [CommandOption("--formulas <JSON>")]
        [Description("Cell formulas as JSON 2D array")]
        public string? Formulas { get; init; }

        [CommandOption("--formulas-file <PATH>")]
        [Description("Read formulas JSON from file")]
        public string? FormulasFile { get; init; }

        [CommandOption("--formats <JSON>")]
        [Description("Number formats as JSON 2D array")]
        public string? Formats { get; init; }

        [CommandOption("--formats-file <PATH>")]
        [Description("Read formats JSON from file")]
        public string? FormatsFile { get; init; }

        [CommandOption("--format-code <CODE>")]
        [Description("Number format code (e.g., #,##0.00)")]
        public string? FormatCode { get; init; }

        // Copy options
        [CommandOption("--source-sheet <NAME>")]
        [Description("Source worksheet for copy")]
        public string? SourceSheet { get; init; }

        [CommandOption("--source-range <ADDRESS>")]
        [Description("Source range for copy")]
        public string? SourceRange { get; init; }

        [CommandOption("--target-sheet <NAME>")]
        [Description("Target worksheet for copy")]
        public string? TargetSheet { get; init; }

        [CommandOption("--target-range <ADDRESS>")]
        [Description("Target range for copy")]
        public string? TargetRange { get; init; }

        // Edit options
        [CommandOption("--shift <DIRECTION>")]
        [Description("Shift direction (Right, Down, Left, Up)")]
        public string? ShiftDirection { get; init; }

        [CommandOption("--search <VALUE>")]
        [Description("Search value for find/replace")]
        public string? SearchValue { get; init; }

        [CommandOption("--replace <VALUE>")]
        [Description("Replace value")]
        public string? ReplaceValue { get; init; }

        [CommandOption("--match-case")]
        [Description("Case-sensitive search")]
        public bool MatchCase { get; init; }

        [CommandOption("--match-entire-cell")]
        [Description("Match entire cell contents")]
        public bool MatchEntireCell { get; init; }

        [CommandOption("--search-formulas")]
        [Description("Search in formulas")]
        public bool SearchFormulas { get; init; }

        [CommandOption("--replace-all")]
        [Description("Replace all occurrences")]
        public bool ReplaceAll { get; init; }

        [CommandOption("--has-headers")]
        [Description("Range has header row")]
        public bool HasHeaders { get; init; }

        [CommandOption("--sort-columns <JSON>")]
        [Description("Sort columns JSON")]
        public string? SortColumnsJson { get; init; }

        // Format options
        [CommandOption("--style-name <NAME>")]
        [Description("Style name")]
        public string? StyleName { get; init; }

        [CommandOption("--font-name <NAME>")]
        [Description("Font name")]
        public string? FontName { get; init; }

        [CommandOption("--font-size <SIZE>")]
        [Description("Font size")]
        public double? FontSize { get; init; }

        [CommandOption("--bold")]
        [Description("Bold text")]
        public bool? Bold { get; init; }

        [CommandOption("--italic")]
        [Description("Italic text")]
        public bool? Italic { get; init; }

        [CommandOption("--underline")]
        [Description("Underline text")]
        public bool? Underline { get; init; }

        [CommandOption("--font-color <COLOR>")]
        [Description("Font color (hex or name)")]
        public string? FontColor { get; init; }

        [CommandOption("--fill-color <COLOR>")]
        [Description("Fill/background color")]
        public string? FillColor { get; init; }

        [CommandOption("--border-style <STYLE>")]
        [Description("Border style")]
        public string? BorderStyle { get; init; }

        [CommandOption("--border-color <COLOR>")]
        [Description("Border color")]
        public string? BorderColor { get; init; }

        [CommandOption("--border-weight <WEIGHT>")]
        [Description("Border weight")]
        public string? BorderWeight { get; init; }

        [CommandOption("--h-align <ALIGNMENT>")]
        [Description("Horizontal alignment")]
        public string? HorizontalAlignment { get; init; }

        [CommandOption("--v-align <ALIGNMENT>")]
        [Description("Vertical alignment")]
        public string? VerticalAlignment { get; init; }

        [CommandOption("--wrap-text")]
        [Description("Wrap text")]
        public bool? WrapText { get; init; }

        [CommandOption("--orientation <DEGREES>")]
        [Description("Text orientation in degrees")]
        public int? Orientation { get; init; }

        // Validation options
        [CommandOption("--validation-type <TYPE>")]
        [Description("Validation type")]
        public string? ValidationType { get; init; }

        [CommandOption("--validation-operator <OP>")]
        [Description("Validation operator")]
        public string? ValidationOperator { get; init; }

        [CommandOption("--formula1 <FORMULA>")]
        [Description("Validation formula 1")]
        public string? Formula1 { get; init; }

        [CommandOption("--formula2 <FORMULA>")]
        [Description("Validation formula 2")]
        public string? Formula2 { get; init; }

        [CommandOption("--show-input-message")]
        [Description("Show input message")]
        public bool? ShowInputMessage { get; init; }

        [CommandOption("--input-title <TITLE>")]
        [Description("Input message title")]
        public string? InputTitle { get; init; }

        [CommandOption("--input-message <MESSAGE>")]
        [Description("Input message")]
        public string? InputMessage { get; init; }

        [CommandOption("--show-error-alert")]
        [Description("Show error alert")]
        public bool? ShowErrorAlert { get; init; }

        [CommandOption("--error-style <STYLE>")]
        [Description("Error style")]
        public string? ErrorStyle { get; init; }

        [CommandOption("--error-title <TITLE>")]
        [Description("Error title")]
        public string? ErrorTitle { get; init; }

        [CommandOption("--error-message <MESSAGE>")]
        [Description("Error message")]
        public string? ErrorMessage { get; init; }

        [CommandOption("--ignore-blank")]
        [Description("Ignore blank cells")]
        public bool? IgnoreBlank { get; init; }

        [CommandOption("--show-dropdown")]
        [Description("Show dropdown for list validation")]
        public bool? ShowDropdown { get; init; }

        // Hyperlink options
        [CommandOption("--url <URL>")]
        [Description("Hyperlink URL")]
        public string? Url { get; init; }

        [CommandOption("--display-text <TEXT>")]
        [Description("Hyperlink display text")]
        public string? DisplayText { get; init; }

        [CommandOption("--tooltip <TEXT>")]
        [Description("Hyperlink tooltip")]
        public string? Tooltip { get; init; }

        // Lock options
        [CommandOption("--locked")]
        [Description("Cell locked state")]
        public bool? Locked { get; init; }
    }
}
