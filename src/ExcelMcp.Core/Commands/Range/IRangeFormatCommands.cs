using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range formatting - fonts, colors, borders, number formats, validation, merge, autofit.
/// Use range command for values/formulas/copy/clear operations.
/// </summary>
[ServiceCategory("rangeformat", "RangeFormat")]
[McpTool("excel_range_format")]
public interface IRangeFormatCommands
{
    // === STYLE OPERATIONS ===

    /// <summary>
    /// Applies built-in Excel cell style to range (recommended for consistency).
    /// Excel COM: Range.Style = styleName
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <param name="styleName">Built-in style name (e.g., "Heading 1", "Good", "Currency"). Use "Normal" to reset.</param>
    [ServiceAction("set-style")]
    void SetStyle(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string styleName);

    /// <summary>
    /// Gets the current built-in style name applied to a range.
    /// Excel COM: Range.Style.Name property
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address</param>
    [ServiceAction("get-style")]
    RangeStyleResult GetStyle(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Applies visual formatting to range (font, fill, border, alignment).
    /// Excel COM: Range.Font, Range.Interior, Range.Borders, Range.HorizontalAlignment, etc.
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address to format</param>
    /// <param name="fontName">Font family name (e.g., "Calibri", "Arial")</param>
    /// <param name="fontSize">Font size in points</param>
    /// <param name="bold">Bold formatting</param>
    /// <param name="italic">Italic formatting</param>
    /// <param name="underline">Underline formatting</param>
    /// <param name="fontColor">Font color (#RRGGBB or color name)</param>
    /// <param name="fillColor">Background fill color (#RRGGBB or color name)</param>
    /// <param name="borderStyle">Border line style</param>
    /// <param name="borderColor">Border color (#RRGGBB or color name)</param>
    /// <param name="borderWeight">Border thickness</param>
    /// <param name="horizontalAlignment">Horizontal text alignment</param>
    /// <param name="verticalAlignment">Vertical text alignment</param>
    /// <param name="wrapText">Wrap text within cell</param>
    /// <param name="orientation">Text rotation angle (-90 to 90)</param>
    /// <remarks>
    /// For consistent, professional formatting, prefer SetStyle with built-in styles.
    /// Use FormatRange only when built-in styles don't meet your needs.
    /// </remarks>
    [ServiceAction("format-range")]
    void FormatRange(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string rangeAddress,
        string? fontName,
        double? fontSize,
        bool? bold,
        bool? italic,
        bool? underline,
        string? fontColor,
        string? fillColor,
        string? borderStyle,
        string? borderColor,
        string? borderWeight,
        string? horizontalAlignment,
        string? verticalAlignment,
        bool? wrapText,
        int? orientation);

    // === VALIDATION OPERATIONS ===

    /// <summary>
    /// Adds data validation rules to range.
    /// Excel COM: Range.Validation.Add()
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address to validate</param>
    /// <param name="validationType">Type of validation (List, WholeNumber, Decimal, Date, etc.)</param>
    /// <param name="validationOperator">Comparison operator (Between, Equal, GreaterThan, etc.)</param>
    /// <param name="formula1">First value/formula for validation</param>
    /// <param name="formula2">Second value for Between/NotBetween operators</param>
    /// <param name="showInputMessage">Show input message when cell selected</param>
    /// <param name="inputTitle">Input message title</param>
    /// <param name="inputMessage">Input message text</param>
    /// <param name="showErrorAlert">Show error alert on invalid entry</param>
    /// <param name="errorStyle">Error alert style (Stop, Warning, Information)</param>
    /// <param name="errorTitle">Error alert title</param>
    /// <param name="errorMessage">Error alert message</param>
    /// <param name="ignoreBlank">Allow blank entries</param>
    /// <param name="showDropdown">Show dropdown for List validation</param>
    [ServiceAction("validate-range")]
    void ValidateRange(
        IExcelBatch batch,
        string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter] string validationType,
        string? validationOperator,
        string? formula1,
        string? formula2,
        bool? showInputMessage,
        string? inputTitle,
        string? inputMessage,
        bool? showErrorAlert,
        string? errorStyle,
        string? errorTitle,
        string? errorMessage,
        bool? ignoreBlank,
        bool? showDropdown);

    /// <summary>
    /// Gets data validation settings from first cell in range.
    /// Excel COM: Range.Validation
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address</param>
    [ServiceAction("get-validation")]
    RangeValidationResult GetValidation(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Removes data validation from range.
    /// Excel COM: Range.Validation.Delete()
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address</param>
    [ServiceAction("remove-validation")]
    void RemoveValidation(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === AUTO-FIT OPERATIONS ===

    /// <summary>
    /// Auto-fits column widths to content.
    /// Excel COM: Range.Columns.AutoFit()
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range defining columns to auto-fit</param>
    [ServiceAction("auto-fit-columns")]
    void AutoFitColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Auto-fits row heights to content.
    /// Excel COM: Range.Rows.AutoFit()
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range defining rows to auto-fit</param>
    [ServiceAction("auto-fit-rows")]
    void AutoFitRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === MERGE OPERATIONS ===

    /// <summary>
    /// Merges cells in range into a single cell.
    /// Excel COM: Range.Merge()
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range to merge</param>
    [ServiceAction("merge-cells")]
    void MergeCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Unmerges previously merged cells.
    /// Excel COM: Range.UnMerge()
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range to unmerge</param>
    [ServiceAction("unmerge-cells")]
    void UnmergeCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Checks if range contains merged cells.
    /// Excel COM: Range.MergeCells
    /// </summary>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range to check</param>
    [ServiceAction("get-merge-info")]
    RangeMergeInfoResult GetMergeInfo(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);
}
