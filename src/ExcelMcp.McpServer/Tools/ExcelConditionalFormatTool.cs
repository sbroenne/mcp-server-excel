using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel conditional formatting operations.
/// </summary>
[McpServerToolType]
public static class ExcelConditionalFormatTool
{
    /// <summary>
    /// Manage Excel conditional formatting rules - visual formatting based on cell values
    /// </summary>
    [McpServerTool(Name = "excel_conditionalformat")]
    [Description(@"Manage Excel conditional formatting - visual formatting based on cell values.

RULE TYPES (non-enum parameter):
- 'cell-value': Format based on cell value comparison (requires operatorType and formula1)
- 'expression': Format based on formula result (requires formula1 only)

OPERATORS (for cell-value type):
- 'equal', 'not-equal', 'greater', 'less', 'greater-equal', 'less-equal'
- 'between', 'not-between' (require both formula1 and formula2)

FORMATTING:
- Interior: interiorColor (#RRGGBB or color index), interiorPattern (solid, gray75, gray50, etc.)
- Font: fontColor (#RRGGBB), fontBold (true/false), fontItalic (true/false)
- Borders: borderStyle (continuous, dash, dot, etc.), borderColor (#RRGGBB)

Example: Highlight cells > 100 in red:
  ruleType='cell-value', operatorType='greater', formula1='100', interiorColor='#FF0000'")]
    public static async Task<string> ExcelConditionalFormat(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        ConditionalFormatAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

        [Description("Worksheet name (required for add-rule and clear-rules actions)")]
        string? sheetName = null,

        [Description("Range address to apply conditional formatting (e.g., 'A1:D10', required for add-rule and clear-rules)")]
        string? rangeAddress = null,

        [Description("Rule type: 'cell-value' or 'expression' (required for add-rule)")]
        string? ruleType = null,

        [Description("Comparison operator: 'equal', 'not-equal', 'greater', 'less', 'greater-equal', 'less-equal', 'between', 'not-between' (required for cell-value rules)")]
        string? operatorType = null,

        [Description("First formula/value for comparison (required for add-rule)")]
        string? formula1 = null,

        [Description("Second formula/value (required for 'between' and 'not-between' operators)")]
        string? formula2 = null,

        [Description("Interior fill color (#RRGGBB hex or color index, e.g., '#FF0000' for red)")]
        string? interiorColor = null,

        [Description("Interior fill pattern: 'solid', 'gray75', 'gray50', 'gray25', 'horizontal', 'vertical', 'down', 'up', 'checker', 'semi-gray75', 'light-horizontal', 'light-vertical', 'light-down', 'light-up', 'grid', 'crisscross', 'gray16', 'gray8'")]
        string? interiorPattern = null,

        [Description("Font color (#RRGGBB hex or color index, e.g., '#0000FF' for blue)")]
        string? fontColor = null,

        [Description("Font bold (true/false)")]
        bool? fontBold = null,

        [Description("Font italic (true/false)")]
        bool? fontItalic = null,

        [Description("Border line style: 'continuous', 'dash', 'dot', 'dash-dot', 'dash-dot-dot', 'slant-dash-dot', 'double'")]
        string? borderStyle = null,

        [Description("Border color (#RRGGBB hex or color index, e.g., '#000000' for black)")]
        string? borderColor = null)
    {
        try
        {
            var conditionalFormattingCommands = new ConditionalFormattingCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                ConditionalFormatAction.AddRule => await AddRuleAsync(
                    conditionalFormattingCommands, sessionId, sheetName, rangeAddress, ruleType, operatorType,
                    formula1, formula2, interiorColor, interiorPattern, fontColor, fontBold, fontItalic,
                    borderStyle, borderColor),
                ConditionalFormatAction.ClearRules => await ClearRulesAsync(
                    conditionalFormattingCommands, sessionId, sheetName, rangeAddress),
                _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed for '{excelPath}': {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static async Task<string> AddRuleAsync(
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.AddRuleAsync(
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
                borderColor));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? "Rule applied. Use excel_range 'get-values' to verify formatting or add more rules to same range."
                : null
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearRulesAsync(
        ConditionalFormattingCommands commands,
        string sessionId,
        string? sheetName,
        string? rangeAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for clear-rules action", nameof(sheetName));
        if (string.IsNullOrEmpty(rangeAddress))
            throw new ArgumentException("rangeAddress is required for clear-rules action", nameof(rangeAddress));

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ClearRulesAsync(batch, sheetName, rangeAddress));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? "All conditional formatting rules removed from range. Cell values unchanged, only formatting cleared."
                : null
        }, ExcelToolsBase.JsonOptions);
    }
}
