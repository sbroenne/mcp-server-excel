using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel conditional formatting operations.
/// </summary>
[McpServerToolType]
public static partial class ExcelConditionalFormatTool
{
    /// <summary>
    /// Conditional formatting - visual rules based on cell values.
    /// TYPES: cell-value (requires operatorType+formula1), expression (formula only).
    /// FORMAT: interiorColor/fontColor #RRGGBB, fontBold/Italic, borderStyle/Color.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="path">Excel file path (.xlsx or .xlsm)</param>
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
    [McpServerTool(Name = "excel_conditionalformat", Title = "Excel Conditional Formatting", Destructive = true)]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelConditionalFormat(
        ConditionalFormatAction action,
        string path,
        string sessionId,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? rangeAddress,
        [DefaultValue(null)] string? ruleType,
        [DefaultValue(null)] string? operatorType,
        [DefaultValue(null)] string? formula1,
        [DefaultValue(null)] string? formula2,
        [DefaultValue(null)] string? interiorColor,
        [DefaultValue(null)] string? interiorPattern,
        [DefaultValue(null)] string? fontColor,
        [DefaultValue(null)] bool? fontBold,
        [DefaultValue(null)] bool? fontItalic,
        [DefaultValue(null)] string? borderStyle,
        [DefaultValue(null)] string? borderColor)
    {
        _ = path; // retained parameter for schema compatibility

        return ExcelToolsBase.ExecuteToolAction(
            "excel_conditionalformat",
            ServiceRegistry.ConditionalFormat.ToActionString(action),
            path,
            () => ServiceRegistry.ConditionalFormat.RouteAction(
                action,
                sessionId,
                ExcelToolsBase.ForwardToServiceFunc,
                sheetName: sheetName,
                rangeAddress: rangeAddress,
                ruleType: ruleType,
                operatorType: operatorType,
                formula1: formula1,
                formula2: formula2,
                interiorColor: interiorColor,
                interiorPattern: interiorPattern,
                fontColor: fontColor,
                fontBold: fontBold,
                fontItalic: fontItalic,
                borderStyle: borderStyle,
                borderColor: borderColor));
    }
}





