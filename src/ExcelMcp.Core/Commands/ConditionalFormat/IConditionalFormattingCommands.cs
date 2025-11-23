using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Commands for managing Excel conditional formatting
/// Excel COM: Range.FormatConditions
/// </summary>
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
    void AddRule(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string ruleType,
        string? operatorType,
        string? formula1,
        string? formula2,
        string? interiorColor = null,
        string? interiorPattern = null,
        string? fontColor = null,
        bool? fontBold = null,
        bool? fontItalic = null,
        string? borderStyle = null,
        string? borderColor = null);

    /// <summary>
    /// Removes all conditional formatting from range
    /// Excel COM: Range.FormatConditions.Delete()
    /// </summary>
    /// <exception cref="InvalidOperationException">Sheet or range not found</exception>
    void ClearRules(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress);
}

