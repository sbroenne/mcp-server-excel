using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range formatting operations: apply styles, set fonts/colors/borders, add data validation, merge cells, auto-fit dimensions.
/// Use range for values/formulas/copy/clear operations.
///
/// STYLES: Use built-in style names like 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent', etc.
/// For consistent, professional formatting, prefer set-style with built-in styles over format-range.
///
/// FONT/COLOR FORMATTING: Specify individual formatting properties:
/// - Colors as hex '#RRGGBB' (e.g., '#FF0000' for red, '#00FF00' for green)
/// - Font sizes as points (e.g., 12, 14, 16)
/// - Alignment: 'left', 'center', 'right' (horizontal), 'top', 'middle', 'bottom' (vertical)
///
/// DATA VALIDATION: Restrict cell input with validation rules:
/// - Types: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'
/// - For list validation, formula1 is the list source (e.g., '=$A$1:$A$10' or '"Option1,Option2,Option3"')
/// - Operators: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'
///
/// MERGE: Combines cells into one. Only top-left cell value is preserved.
/// </summary>
[ServiceCategory("rangeformat", "RangeFormat")]
[McpTool("excel_range_format", Title = "Excel Range Format Operations", Destructive = true, Category = "data",
    Description = "Range formatting: styles, fonts/colors/borders, data validation, merge cells, auto-fit. STYLES: Built-in names (Heading 1, Good, Bad, Currency, Percent). COLORS: Hex #RRGGBB (e.g., #FF0000 red, #00FF00 green). FONT: size in points, alignment horizontal (left/center/right) and vertical (top/middle/bottom). DATA VALIDATION: Types list, whole, decimal, date, time, textLength, custom. For list: validationFormula1 is source (=$A$1:$A$10 or Option1,Option2). Operators: between, equal, greaterThan, etc. MERGE: Only top-left cell value preserved.")]
public interface IRangeFormatCommands
{
    // === STYLE OPERATIONS ===

    /// <summary>
    /// Applies built-in Excel cell style to range (recommended for consistency).
    /// Excel COM: Range.Style = styleName
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    /// <param name="styleName">Built-in or custom style name (e.g., 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent'). Use 'Normal' to reset.</param>
    [ServiceAction("set-style")]
    void SetStyle(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string styleName);

    /// <summary>
    /// Gets the current built-in style name applied to a range.
    /// Excel COM: Range.Style.Name property
    /// </summary>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    [ServiceAction("get-style")]
    RangeStyleResult GetStyle(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Applies visual formatting to range (font, fill, border, alignment).
    /// Excel COM: Range.Font, Range.Interior, Range.Borders, Range.HorizontalAlignment, etc.
    /// </summary>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to format (e.g., 'A1:D10')</param>
    /// <param name="fontName">Font family name (e.g., 'Arial', 'Calibri', 'Times New Roman')</param>
    /// <param name="fontSize">Font size in points (e.g., 10, 11, 12, 14, 16)</param>
    /// <param name="bold">Whether to apply bold formatting</param>
    /// <param name="italic">Whether to apply italic formatting</param>
    /// <param name="underline">Whether to apply underline formatting</param>
    /// <param name="fontColor">Font (foreground) color as hex '#RRGGBB' (e.g., '#FF0000' for red)</param>
    /// <param name="fillColor">Cell fill (background) color as hex '#RRGGBB' (e.g., '#FFFF00' for yellow)</param>
    /// <param name="borderStyle">Border line style (e.g., 'thin', 'medium', 'thick', 'dashed', 'dotted')</param>
    /// <param name="borderColor">Border color as hex '#RRGGBB'</param>
    /// <param name="borderWeight">Border weight (e.g., 'hairline', 'thin', 'medium', 'thick')</param>
    /// <param name="horizontalAlignment">Horizontal text alignment: 'left', 'center', 'right', 'justify', 'fill'</param>
    /// <param name="verticalAlignment">Vertical text alignment: 'top', 'middle', 'bottom', 'justify'</param>
    /// <param name="wrapText">Whether to wrap text within cells</param>
    /// <param name="orientation">Text rotation in degrees (-90 to 90, or 255 for vertical)</param>
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
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to validate (e.g., 'A1:D10')</param>
    /// <param name="validationType">Data validation type: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'</param>
    /// <param name="validationOperator">Validation comparison operator: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'</param>
    /// <param name="formula1">First validation formula/value - for list validation use range '=$A$1:$A$10' or inline '"A,B,C"'</param>
    /// <param name="formula2">Second validation formula/value - required only for 'between' and 'notBetween' operators</param>
    /// <param name="showInputMessage">Whether to show input message when cell is selected (default: false)</param>
    /// <param name="inputTitle">Title for the input message popup</param>
    /// <param name="inputMessage">Text for the input message popup</param>
    /// <param name="showErrorAlert">Whether to show error alert on invalid input (default: true)</param>
    /// <param name="errorStyle">Error alert style: 'stop' (prevents entry), 'warning' (allows override), 'information' (allows entry)</param>
    /// <param name="errorTitle">Title for the error alert popup</param>
    /// <param name="errorMessage">Text for the error alert popup</param>
    /// <param name="ignoreBlank">Whether to allow blank cells in validation (default: true)</param>
    /// <param name="showDropdown">Whether to show dropdown arrow for list validation (default: true)</param>
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
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    [ServiceAction("get-validation")]
    RangeValidationResult GetValidation(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Removes data validation from range.
    /// Excel COM: Range.Validation.Delete()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    [ServiceAction("remove-validation")]
    void RemoveValidation(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === AUTO-FIT OPERATIONS ===

    /// <summary>
    /// Auto-fits column widths to content.
    /// Excel COM: Range.Columns.AutoFit()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Column range to auto-fit (e.g., 'A:D' or 'A1:D100')</param>
    [ServiceAction("auto-fit-columns")]
    void AutoFitColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Auto-fits row heights to content.
    /// Excel COM: Range.Rows.AutoFit()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row range to auto-fit (e.g., '1:10' or 'A1:D100')</param>
    [ServiceAction("auto-fit-rows")]
    void AutoFitRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === MERGE OPERATIONS ===

    /// <summary>
    /// Merges cells in range into a single cell.
    /// Excel COM: Range.Merge()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range to merge into a single cell (e.g., 'A1:D1')</param>
    [ServiceAction("merge-cells")]
    void MergeCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Unmerges previously merged cells.
    /// Excel COM: Range.UnMerge()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range to unmerge (e.g., 'A1:D1')</param>
    [ServiceAction("unmerge-cells")]
    void UnmergeCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Checks if range contains merged cells.
    /// Excel COM: Range.MergeCells
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range to check for merged cells (e.g., 'A1:D10')</param>
    [ServiceAction("get-merge-info")]
    RangeMergeInfoResult GetMergeInfo(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);
}
