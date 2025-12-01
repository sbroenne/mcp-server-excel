using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel conditional formatting operations.
/// </summary>
[McpServerToolType]
public static partial class ExcelConditionalFormatTool
{
    /// <summary>
    /// Manage Excel conditional formatting - visual formatting based on cell values.
    /// RULE TYPES: 'cell-value' (format based on cell value comparison, requires operatorType and formula1) or 'expression' (format based on formula result, requires formula1 only).
    /// OPERATORS for cell-value type: 'equal', 'not-equal', 'greater', 'less', 'greater-equal', 'less-equal', 'between', 'not-between' (between/not-between require both formula1 and formula2).
    /// FORMATTING: Interior (interiorColor #RRGGBB, interiorPattern), Font (fontColor #RRGGBB, fontBold, fontItalic), Borders (borderStyle, borderColor).
    /// Example: Highlight cells greater than 100 in red: ruleType='cell-value', operatorType='greater', formula1='100', interiorColor='#FF0000'.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Excel file path (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="sheetName">Worksheet name (required for add-rule and clear-rules actions)</param>
    /// <param name="rangeAddress">Range address to apply conditional formatting (e.g., 'A1:D10', required for add-rule and clear-rules)</param>
    /// <param name="ruleType">Rule type: 'cell-value' or 'expression' (required for add-rule)</param>
    /// <param name="operatorType">Comparison operator: 'equal', 'not-equal', 'greater', 'less', 'greater-equal', 'less-equal', 'between', 'not-between' (required for cell-value rules)</param>
    /// <param name="formula1">First formula/value for comparison (required for add-rule)</param>
    /// <param name="formula2">Second formula/value (required for 'between' and 'not-between' operators)</param>
    /// <param name="interiorColor">Interior fill color (#RRGGBB hex or color index, e.g., '#FF0000' for red)</param>
    /// <param name="interiorPattern">Interior fill pattern: 'solid', 'gray75', 'gray50', 'gray25', 'horizontal', 'vertical', 'down', 'up', 'checker', 'semi-gray75', 'light-horizontal', 'light-vertical', 'light-down', 'light-up', 'grid', 'crisscross', 'gray16', 'gray8'</param>
    /// <param name="fontColor">Font color (#RRGGBB hex or color index, e.g., '#0000FF' for blue)</param>
    /// <param name="fontBold">Font bold (true/false)</param>
    /// <param name="fontItalic">Font italic (true/false)</param>
    /// <param name="borderStyle">Border line style: 'continuous', 'dash', 'dot', 'dash-dot', 'dash-dot-dot', 'slant-dash-dot', 'double'</param>
    /// <param name="borderColor">Border color (#RRGGBB hex or color index, e.g., '#000000' for black)</param>
    [McpServerTool(Name = "excel_conditionalformat")]
    [McpMeta("category", "structure")]
    public static partial string ExcelConditionalFormat(
        ConditionalFormatAction action,
        string excelPath,
        string sessionId,
        string? sheetName,
        string? rangeAddress,
        string? ruleType,
        string? operatorType,
        string? formula1,
        string? formula2,
        string? interiorColor,
        string? interiorPattern,
        string? fontColor,
        bool? fontBold,
        bool? fontItalic,
        string? borderStyle,
        string? borderColor)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_conditionalformat",
            action.ToActionString(),
            excelPath,
            () =>
            {
                var conditionalFormattingCommands = new ConditionalFormattingCommands();

                // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
                return action switch
                {
                    ConditionalFormatAction.AddRule => AddRuleAsync(
                        conditionalFormattingCommands, sessionId, sheetName, rangeAddress, ruleType, operatorType,
                        formula1, formula2, interiorColor, interiorPattern, fontColor, fontBold, fontItalic,
                        borderStyle, borderColor),
                    ConditionalFormatAction.ClearRules => ClearRulesAsync(
                        conditionalFormattingCommands, sessionId, sheetName, rangeAddress),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string AddRuleAsync(
        ConditionalFormattingCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress,
        string? ruleType,
        string? operatorType,
        string? formula1,
        string? formula2,
        string? interiorColor,
        string? interiorPattern,
        string? fontColor,
        bool? fontBold,
        bool? fontItalic,
        string? borderStyle,
        string? borderColor)
    {
        // Validate required parameters
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for add-rule action", nameof(sheetName));
        if (string.IsNullOrEmpty(rangeAddress))
            throw new ArgumentException("rangeAddress is required for add-rule action", nameof(rangeAddress));
        if (string.IsNullOrEmpty(ruleType))
            throw new ArgumentException("ruleType is required for add-rule action", nameof(ruleType));
        if (string.IsNullOrEmpty(formula1))
            throw new ArgumentException("formula1 is required for add-rule action", nameof(formula1));

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.AddRule(
                    batch,
                    sheetName,
                    rangeAddress,
                    ruleType,
                    operatorType,
                    formula1,
                    formula2,
                    interiorColor,
                    interiorPattern,
                    fontColor,
                    fontBold,
                    fontItalic,
                    borderStyle,
                    borderColor);
                return 0; // Dummy return value for WithSession
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = "Conditional formatting rule added successfully"
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ClearRulesAsync(
        ConditionalFormattingCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for clear-rules action", nameof(sheetName));
        if (string.IsNullOrEmpty(rangeAddress))
            throw new ArgumentException("rangeAddress is required for clear-rules action", nameof(rangeAddress));

        ExcelToolsBase.WithSession(
            sessionId,
            batch =>
            {
                commands.ClearRules(batch, sheetName, rangeAddress);
                return 0; // Dummy return value for WithSession
            });

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = "All conditional formatting rules removed from range"
        }, ExcelToolsBase.JsonOptions);
    }
}

