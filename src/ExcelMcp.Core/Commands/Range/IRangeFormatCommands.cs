using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Excel range formatting operations - styling, validation, merge, autofit.
/// Use IRangeCommands for values/formulas/copy/clear operations.
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
    [ServiceAction("get-style")]
    RangeStyleResult GetStyle(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Applies visual formatting to range (font, fill, border, alignment).
    /// Excel COM: Range.Font, Range.Interior, Range.Borders, Range.HorizontalAlignment, etc.
    /// </summary>
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
    [ServiceAction("get-validation")]
    RangeValidationResult GetValidation(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Removes data validation from range.
    /// Excel COM: Range.Validation.Delete()
    /// </summary>
    [ServiceAction("remove-validation")]
    void RemoveValidation(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === AUTO-FIT OPERATIONS ===

    /// <summary>
    /// Auto-fits column widths to content.
    /// Excel COM: Range.Columns.AutoFit()
    /// </summary>
    [ServiceAction("auto-fit-columns")]
    void AutoFitColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Auto-fits row heights to content.
    /// Excel COM: Range.Rows.AutoFit()
    /// </summary>
    [ServiceAction("auto-fit-rows")]
    void AutoFitRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === MERGE OPERATIONS ===

    /// <summary>
    /// Merges cells in range into a single cell.
    /// Excel COM: Range.Merge()
    /// </summary>
    [ServiceAction("merge-cells")]
    void MergeCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Unmerges previously merged cells.
    /// Excel COM: Range.UnMerge()
    /// </summary>
    [ServiceAction("unmerge-cells")]
    void UnmergeCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Checks if range contains merged cells.
    /// Excel COM: Range.MergeCells
    /// </summary>
    [ServiceAction("get-merge-info")]
    RangeMergeInfoResult GetMergeInfo(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);
}
