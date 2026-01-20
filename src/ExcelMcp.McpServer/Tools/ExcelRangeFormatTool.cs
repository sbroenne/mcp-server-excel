using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel range formatting - styles, validation, merge, autofit.
/// </summary>
[McpServerToolType]
public static partial class ExcelRangeFormatTool
{
    /// <summary>
    /// Range formatting operations: apply styles, set fonts/colors/borders, add data validation, merge cells, auto-fit dimensions.
    ///
    /// STYLES: Use built-in style names like 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent', etc.
    ///
    /// FONT/COLOR FORMATTING: Specify individual formatting properties:
    /// - Colors as hex '#RRGGBB' (e.g., '#FF0000' for red, '#00FF00' for green)
    /// - Font sizes as points (e.g., 12, 14, 16)
    /// - Alignment: 'left', 'center', 'right' (horizontal), 'top', 'middle', 'bottom' (vertical)
    ///
    /// DATA VALIDATION: Restrict cell input with validation rules:
    /// - Types: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'
    /// - For list validation, validationFormula1 is the list source (e.g., '=$A$1:$A$10' or '"Option1,Option2,Option3"')
    /// - Operators: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'
    ///
    /// MERGE: Combines cells into one. Only top-left cell value is preserved.
    /// </summary>
    /// <param name="action">The range format operation to perform</param>
    /// <param name="excelPath">Full path to the Excel workbook file (e.g., 'C:\Reports\Sales.xlsx')</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action - required for all operations</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10', 'B:D' for columns)</param>
    /// <param name="styleName">Built-in or custom style name (e.g., 'Heading 1', 'Good', 'Currency', 'Percent')</param>
    /// <param name="fontName">Font family name (e.g., 'Arial', 'Calibri', 'Times New Roman')</param>
    /// <param name="fontSize">Font size in points (e.g., 10, 11, 12, 14, 16)</param>
    /// <param name="bold">Whether to apply bold formatting</param>
    /// <param name="italic">Whether to apply italic formatting</param>
    /// <param name="underline">Whether to apply underline formatting</param>
    /// <param name="fontColor">Font (foreground) color as hex '#RRGGBB' (e.g., '#FF0000' for red)</param>
    /// <param name="backgroundColor">Cell fill (background) color as hex '#RRGGBB' (e.g., '#FFFF00' for yellow)</param>
    /// <param name="borderStyle">Border line style (e.g., 'thin', 'medium', 'thick', 'dashed', 'dotted')</param>
    /// <param name="borderColor">Border color as hex '#RRGGBB'</param>
    /// <param name="borderWeight">Border weight (e.g., 'hairline', 'thin', 'medium', 'thick')</param>
    /// <param name="horizontalAlignment">Horizontal text alignment: 'left', 'center', 'right', 'justify', 'fill'</param>
    /// <param name="verticalAlignment">Vertical text alignment: 'top', 'middle', 'bottom', 'justify'</param>
    /// <param name="wrapText">Whether to wrap text within cells</param>
    /// <param name="textOrientation">Text rotation in degrees (-90 to 90, or 255 for vertical)</param>
    /// <param name="validationType">Data validation type: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'</param>
    /// <param name="validationOperator">Validation comparison operator: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'</param>
    /// <param name="validationFormula1">First validation formula/value - for list validation use range '=$A$1:$A$10' or inline '"A,B,C"'</param>
    /// <param name="validationFormula2">Second validation formula/value - required only for 'between' and 'notBetween' operators</param>
    /// <param name="showInputMessage">Whether to show input message when cell is selected (default: false)</param>
    /// <param name="inputMessageTitle">Title for the input message popup</param>
    /// <param name="inputMessageText">Text for the input message popup</param>
    /// <param name="showErrorAlert">Whether to show error alert on invalid input (default: true)</param>
    /// <param name="errorAlertStyle">Error alert style: 'stop' (prevents entry), 'warning' (allows override), 'information' (allows entry)</param>
    /// <param name="errorAlertTitle">Title for the error alert popup</param>
    /// <param name="errorAlertMessage">Text for the error alert popup</param>
    /// <param name="ignoreBlankCells">Whether to allow blank cells in validation (default: true)</param>
    /// <param name="showDropdownList">Whether to show dropdown arrow for list validation (default: true)</param>
    [McpServerTool(Name = "excel_range_format", Title = "Excel Range Format Operations", Destructive = true)]
    [McpMeta("category", "data")]
    [McpMeta("requiresSession", true)]
    public static partial string RangeFormat(
        RangeFormatAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] string? styleName,
        [DefaultValue(null)] string? fontName,
        [DefaultValue(null)] double? fontSize,
        [DefaultValue(null)] bool? bold,
        [DefaultValue(null)] bool? italic,
        [DefaultValue(null)] bool? underline,
        [DefaultValue(null)] string? fontColor,
        [DefaultValue(null)] string? backgroundColor,
        [DefaultValue(null)] string? borderStyle,
        [DefaultValue(null)] string? borderColor,
        [DefaultValue(null)] string? borderWeight,
        [DefaultValue(null)] string? horizontalAlignment,
        [DefaultValue(null)] string? verticalAlignment,
        [DefaultValue(null)] bool? wrapText,
        [DefaultValue(null)] int? textOrientation,
        [DefaultValue(null)] string? validationType,
        [DefaultValue(null)] string? validationOperator,
        [DefaultValue(null)] string? validationFormula1,
        [DefaultValue(null)] string? validationFormula2,
        [DefaultValue(null)] bool? showInputMessage,
        [DefaultValue(null)] string? inputMessageTitle,
        [DefaultValue(null)] string? inputMessageText,
        [DefaultValue(null)] bool? showErrorAlert,
        [DefaultValue(null)] string? errorAlertStyle,
        [DefaultValue(null)] string? errorAlertTitle,
        [DefaultValue(null)] string? errorAlertMessage,
        [DefaultValue(null)] bool? ignoreBlankCells,
        [DefaultValue(null)] bool? showDropdownList)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_range_format",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var rangeCommands = new RangeCommands();

                return action switch
                {
                    RangeFormatAction.GetStyle => GetStyleAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeFormatAction.SetStyle => SetStyleAction(rangeCommands, sessionId, sheetName, rangeAddress, styleName),
                    RangeFormatAction.FormatRange => FormatRangeAction(rangeCommands, sessionId, sheetName, rangeAddress, fontName, fontSize, bold, italic, underline, fontColor, backgroundColor, borderStyle, borderColor, borderWeight, horizontalAlignment, verticalAlignment, wrapText, textOrientation),
                    RangeFormatAction.ValidateRange => ValidateRangeAction(rangeCommands, sessionId, sheetName, rangeAddress, validationType, validationOperator, validationFormula1, validationFormula2, showInputMessage, inputMessageTitle, inputMessageText, showErrorAlert, errorAlertStyle, errorAlertTitle, errorAlertMessage, ignoreBlankCells, showDropdownList),
                    RangeFormatAction.GetValidation => GetValidationAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeFormatAction.RemoveValidation => RemoveValidationAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeFormatAction.AutoFitColumns => AutoFitColumnsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeFormatAction.AutoFitRows => AutoFitRowsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeFormatAction.MergeCells => MergeCellsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeFormatAction.UnmergeCells => UnmergeCellsAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    RangeFormatAction.GetMergeInfo => GetMergeInfoAction(rangeCommands, sessionId, sheetName, rangeAddress),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string GetStyleAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-style");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetStyle(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            Success = true,
            sheetName,
            rangeAddress,
            styleName = result.StyleName,
            isBuiltInStyle = result.IsBuiltInStyle,
            styleDescription = result.StyleDescription
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetStyleAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress, string? styleName)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "set-style");
        if (string.IsNullOrEmpty(styleName))
            ExcelToolsBase.ThrowMissingParameter("styleName", "set-style");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.SetStyle(batch, sheetName ?? "", rangeAddress!, styleName!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string FormatRangeAction(
        RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress,
        string? fontName, double? fontSize, bool? bold, bool? italic, bool? underline,
        string? fontColor, string? backgroundColor, string? borderStyle, string? borderColor, string? borderWeight,
        string? horizontalAlignment, string? verticalAlignment, bool? wrapText, int? textOrientation)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "format-range");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.FormatRange(batch, sheetName ?? "", rangeAddress!, fontName, fontSize, bold, italic, underline, fontColor, backgroundColor, borderStyle, borderColor, borderWeight, horizontalAlignment, verticalAlignment, wrapText, textOrientation);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string ValidateRangeAction(
        RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress,
        string? validationType, string? validationOperator, string? validationFormula1, string? validationFormula2,
        bool? showInputMessage, string? inputMessageTitle, string? inputMessageText,
        bool? showErrorAlert, string? errorAlertStyle, string? errorAlertTitle, string? errorAlertMessage,
        bool? ignoreBlankCells, bool? showDropdownList)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "validate-range");
        if (string.IsNullOrEmpty(validationType))
            ExcelToolsBase.ThrowMissingParameter("validationType", "validate-range");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.ValidateRange(batch, sheetName ?? "", rangeAddress!, validationType!, validationOperator, validationFormula1, validationFormula2, showInputMessage, inputMessageTitle, inputMessageText, showErrorAlert, errorAlertStyle, errorAlertTitle, errorAlertMessage, ignoreBlankCells, showDropdownList);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string GetValidationAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-validation");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetValidation(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            validationType = ((dynamic)result).ValidationType,
            validationOperator = ((dynamic)result).ValidationOperator,
            validationFormula1 = ((dynamic)result).Formula1,
            validationFormula2 = ((dynamic)result).Formula2,
            showInputMessage = ((dynamic)result).ShowInputMessage,
            inputMessageTitle = ((dynamic)result).InputTitle,
            inputMessageText = ((dynamic)result).InputMessage,
            showErrorAlert = ((dynamic)result).ShowErrorAlert,
            errorAlertStyle = ((dynamic)result).ErrorStyle,
            errorAlertTitle = ((dynamic)result).ErrorTitle,
            errorAlertMessage = ((dynamic)result).ErrorMessage,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RemoveValidationAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "remove-validation");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.RemoveValidation(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string AutoFitColumnsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "auto-fit-columns");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.AutoFitColumns(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string AutoFitRowsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "auto-fit-rows");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.AutoFitRows(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string MergeCellsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "merge-cells");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.MergeCells(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string UnmergeCellsAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "unmerge-cells");

        ExcelToolsBase.WithSession<object?>(
            sessionId,
            batch =>
            {
                commands.UnmergeCells(batch, sheetName ?? "", rangeAddress!);
                return null;
            });

        return JsonSerializer.Serialize(new { Success = true }, ExcelToolsBase.JsonOptions);
    }

    private static string GetMergeInfoAction(RangeCommands commands, string sessionId, string? sheetName, string? rangeAddress)
    {
        if (string.IsNullOrEmpty(rangeAddress))
            ExcelToolsBase.ThrowMissingParameter("rangeAddress", "get-merge-info");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.GetMergeInfo(batch, sheetName ?? "", rangeAddress!));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            isMerged = ((dynamic)result).IsMerged,
            mergeAddress = ((dynamic)result).MergeAddress,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}
