using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Conditional formatting - visual rules based on cell values.
/// TYPES: cellValue (requires operatorType+formula1), expression (formula only). Both camelCase and kebab-case accepted.
/// FORMAT: interiorColor/fontColor as #RRGGBB, fontBold/Italic, borderStyle/Color.
///
/// OPERATORS: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween.
/// For 'between' and 'notBetween', both formula1 and formula2 are required.
/// </summary>
[ServiceCategory("conditionalformat", "ConditionalFormat")]
[McpTool("conditionalformat", Title = "Conditional Formatting", Destructive = true, Category = "structure",
    Description = "Conditional formatting - visual rules based on cell values. TYPES: cellValue (accepts both camelCase and kebab-case, e.g. cell-value), expression. For cellValue: requires operatorType + formula1. FORMAT: interiorColor/fontColor as #RRGGBB hex, fontBold/fontItalic booleans, borderStyle/borderColor.")]
public interface IConditionalFormattingCommands
{
    /// <summary>
    /// Adds conditional formatting rule to range with full format control
    /// Excel COM: Range.FormatConditions.Add(), FormatCondition.Interior/Font/Borders
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <param name="rangeAddress">Range address (A1 notation or named range)</param>
    /// <param name="ruleType">Rule type: cellValue (or cell-value), expression, colorScale, dataBar, top10, iconSet, uniqueValues, blanksCondition, timePeriod, aboveAverage. Both camelCase and kebab-case accepted.</param>
    /// <param name="operatorType">XlFormatConditionOperator: equal, notEqual, greater, less, greaterEqual, lessEqual, between, notBetween</param>
    /// <param name="formula1">First formula/value for condition</param>
    /// <param name="formula2">Second formula/value (for between/notBetween)</param>
    /// <param name="interiorColor">Fill color (#RRGGBB or color index)</param>
    /// <param name="interiorPattern">Interior pattern (1=Solid, -4142=None, 9=Gray50, etc.)</param>
    /// <param name="fontColor">Font color (#RRGGBB or color index)</param>
    /// <param name="fontBold">Bold font</param>
    /// <param name="fontItalic">Italic font</param>
    /// <param name="borderStyle">Border style: none, continuous, dash, dot, etc.</param>
    /// <param name="borderColor">Border color (#RRGGBB or color index)</param>
    /// <exception cref="InvalidOperationException">Sheet or range not found</exception>
    /// <exception cref="ArgumentException">Invalid rule type, operator, color, or format value</exception>
    [ServiceAction("add-rule")]
    OperationResult AddRule(
        IExcelBatch batch,
        [RequiredParameter, FromString("sheetName")] string sheetName,
        [RequiredParameter, FromString("rangeAddress")] string rangeAddress,
        [RequiredParameter, FromString("ruleType")] string ruleType,
        [FromString("operatorType")] string? operatorType,
        [FromString("formula1")] string? formula1,
        [FromString("formula2")] string? formula2,
        [FromString("interiorColor")] string? interiorColor = null,
        [FromString("interiorPattern")] string? interiorPattern = null,
        [FromString("fontColor")] string? fontColor = null,
        [FromString("fontBold")] bool? fontBold = null,
        [FromString("fontItalic")] bool? fontItalic = null,
        [FromString("borderStyle")] string? borderStyle = null,
        [FromString("borderColor")] string? borderColor = null);

    /// <summary>
    /// Removes all conditional formatting from range
    /// Excel COM: Range.FormatConditions.Delete()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="rangeAddress">Range address to clear rules from (e.g., A1:D10)</param>
    /// <exception cref="InvalidOperationException">Sheet or range not found</exception>
    [ServiceAction("clear-rules")]
    OperationResult ClearRules(
        IExcelBatch batch,
        [RequiredParameter, FromString("sheetName")] string sheetName,
        [RequiredParameter, FromString("rangeAddress")] string rangeAddress);

    /// <summary>
    /// Lists existing conditional formatting rules for a range.
    /// Excel COM: Range.FormatConditions enumeration.
    /// Returns rule type, operator, formulas, applies-to range, priority, and formatting
    /// (interior/font/borders). Colors are returned as #RRGGBB hex strings. Rules are returned
    /// in priority order. Rule types that do not use an operator or formatting (e.g. colorScale,
    /// dataBar, iconSet) return only their type with null formatting fields.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <param name="rangeAddress">Range address to read rules from (e.g., A1:G41)</param>
    /// <exception cref="InvalidOperationException">Sheet or range not found</exception>
    [ServiceAction("list-rules")]
    ConditionalFormatListResult ListRules(
        IExcelBatch batch,
        [RequiredParameter, FromString("sheetName")] string sheetName,
        [RequiredParameter, FromString("rangeAddress")] string rangeAddress);

    /// <summary>
    /// Lists all conditional formatting rules on an entire worksheet.
    /// Excel COM: Worksheet.Cells.FormatConditions enumeration.
    /// Each rule includes its applies-to range so rules can be grouped by range. Colors are
    /// returned as #RRGGBB hex strings and rules are returned in priority order.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <exception cref="InvalidOperationException">Sheet not found</exception>
    [ServiceAction("list-worksheet-rules")]
    ConditionalFormatListResult ListWorksheetRules(
        IExcelBatch batch,
        [RequiredParameter, FromString("sheetName")] string sheetName);
}



