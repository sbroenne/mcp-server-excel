using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Commands for managing Excel conditional formatting
/// Excel COM: Range.FormatConditions
/// </summary>
[ServiceCategory("conditionalformat", "ConditionalFormat")]
[McpTool("excel_conditionalformat")]
public interface IConditionalFormattingCommands
{
    /// <summary>
    /// Adds conditional formatting rule to range with full format control
    /// Excel COM: Range.FormatConditions.Add(), FormatCondition.Interior/Font/Borders
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <param name="rangeAddress">Range address (A1 notation or named range)</param>
    /// <param name="ruleType">Rule type: cellValue, expression</param>
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
    void AddRule(
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
    /// <exception cref="InvalidOperationException">Sheet or range not found</exception>
    [ServiceAction("clear-rules")]
    void ClearRules(
        IExcelBatch batch,
        [RequiredParameter, FromString("sheetName")] string sheetName,
        [RequiredParameter, FromString("rangeAddress")] string rangeAddress);
}



