using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range formatting operations: apply styles, set fonts/colors/borders, add data validation, merge cells, auto-fit dimensions.
/// Use range tool for values/formulas/copy/clear operations.
///
/// set-style: Apply a named Excel style (Heading 1, Good, Bad, Neutral, Normal).
/// Best for semantic status labels (Good/Bad/Neutral have fill colours and are theme-aware) and document hierarchy (Heading 1/2/3).
/// NOTE: Heading styles do NOT apply a fill colour — use format when you need a coloured header row.
///
/// format: Apply one typed formatOptions payload to one or more rangeAddresses.
/// Includes independent borders, theme colors/tints, font settings, and indentation.
/// All target ranges are validated before formatting begins. If any target range is invalid, nothing is formatted.
///
/// COLORS: Hex '#RRGGBB' (e.g., '#FF0000' for red, '#00FF00' for green)
/// FONT: size in points (e.g., 12, 14, 16), alignment: 'left', 'center', 'right' / 'top', 'middle', 'bottom'
///
/// DATA VALIDATION: Restrict cell input with validation rules:
/// - Types: 'list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'
/// - For list validation, formula1 is the list source (e.g., '=$A$1:$A$10' or '"Option1,Option2,Option3"')
/// - Operators: 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'
///
/// MERGE: Combines cells into one. Only top-left cell value is preserved.
/// </summary>
[ServiceCategory("rangeformat", "RangeFormat")]
[McpTool("range_format", Title = "Range Format Operations", Destructive = true, Category = "data",
    Description = "Range formatting: styles, custom visual formatting, data validation, merge, auto-fit. " +
        "get-format: Read every requested cell's stored, displayed (including conditional formatting), or both formatting snapshots; no preview limit and no selection changes. " +
        "get-visibility: Read every unique intersecting whole row or column, native current size, outline level and worksheet AutoFilter context; hidden cause is undetermined. " +
        "set-visibility: Required axis rows/columns and hidden true/false. Preserve stored dimensions; do not remove filter criteria or groups. Disjoint gaps remain unchanged. " +
        "set-style: Named styles (Good/Bad/Neutral have fills and are theme-aware; Heading 1/2/3 for document hierarchy; Normal to reset). " +
        "NOTE: Heading styles do NOT include a fill colour — use format for coloured header rows. " +
        "format: One format_options JSON object for one or more range_addresses. Includes independent edge/inside/diagonal borders, font settings, theme colors/tints, indentation, alignment and number format. Omitted settings preserve existing state. " +
        "All targets and known invalid options are checked before mutation. Native failures do not promise rollback. Fixed RGB and theme colors are mutually exclusive for each component. " +
        "COLORS: Hex #RRGGBB. FONT: size in points, alignment left/center/right, top/middle/bottom. " +
        "DATA VALIDATION: Types list/whole/decimal/date/time/textLength/custom. For list: formula1 is source (=$A$1:$A$10 or \"A,B,C\"). " +
        "MERGE: Only top-left cell value preserved. " +
        "TABLES: For Excel Table visual styling use table(action:'set-style') — do not apply range_format to table header or data rows, table style manages all table formatting. " +
        "PIVOTTABLES: Do not apply range_format to PivotTable cells — formatting is overwritten on the next refresh.")]
public interface IRangeFormatCommands
{
    /// <summary>
    /// Reads every unique whole row or column intersecting the exact requested scope.
    /// Includes Hidden, native current size, outline level, and worksheet AutoFilter context.
    /// Hidden cause is undetermined: COM has no reliable flag separating manual hiding,
    /// filtering, zero size, or collapsed groups. Does not change selection or activation.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name; empty for named ranges</param>
    /// <param name="rangeAddress">Exact scope; whole intersecting dimensions are returned without a cap</param>
    /// <param name="axis">rows or columns</param>
    [ServiceAction("get-visibility")]
    RangeVisibilityResult GetVisibility(IExcelBatch batch, [AllowEmptyString] string sheetName,
        [RequiredParameter] string rangeAddress, [RequiredParameter][FromString] VisibilityAxis axis);

    /// <summary>
    /// Sets native Hidden for every whole row or column intersecting the exact scope,
    /// preserving Excel's stored dimensions. Disjoint gaps remain unchanged. Does not
    /// remove filter criteria or outline groups; they can affect visibility again.
    /// Does not bypass protection or promise rollback after a native failure.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name; empty for named ranges</param>
    /// <param name="rangeAddress">Exact scope selecting whole intersecting rows or columns</param>
    /// <param name="axis">rows or columns</param>
    /// <param name="hidden">Required true to hide or false to show; native stored dimensions are retained</param>
    [ServiceAction("set-visibility")]
    OperationResult SetVisibility(IExcelBatch batch, [AllowEmptyString] string sheetName,
        [RequiredParameter] string rangeAddress, [RequiredParameter][FromString] VisibilityAxis axis,
        [RequiredParameter] bool hidden);

    /// <summary>
    /// Reads complete per-cell formatting in the exact requested scope.
    /// Stored reads report workbook formatting; displayed reads include conditional
    /// formatting. Both returns the two snapshots without changing cells or selection.
    /// Mixed rich-text properties are identified explicitly, not replaced with defaults.
    /// There is no preview limit. Use get-style only when the style name is sufficient.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name, or empty string for a named range</param>
    /// <param name="rangeAddress">Exact range or named range to inspect completely</param>
    /// <param name="view">stored (default), displayed, or both formatting snapshots</param>
    [ServiceAction("get-format")]
    RangeFormatReadResult GetFormat(IExcelBatch batch, [AllowEmptyString] string sheetName,
        [RequiredParameter] string rangeAddress, [FromString] FormatView view = FormatView.Stored);

    // === STYLE OPERATIONS ===

    /// <summary>
    /// Applies an existing built-in or custom Excel cell style to a range.
    /// Excel COM: Range.Style = styleName
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    /// <param name="styleName">Built-in or custom style name (e.g., 'Heading 1', 'Good', 'Bad', 'Currency', 'Percent'). Use 'Normal' to reset.</param>
    [ServiceAction("set-style")]
    OperationResult SetStyle(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string styleName);

    /// <summary>
    /// Gets the native cell style applied to a range, including built-in/custom status.
    /// Excel COM: Range.Style.Name property
    /// When Excel reports no single style for a mixed-style range, returns the
    /// existing Normal fallback without changing individual cell styles. This does
    /// not mean every cell uses Normal; use get-format to inspect each cell's distinct style.
    /// Native style-read failures remain errors, not successful Normal results.
    /// </summary>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    [ServiceAction("get-style")]
    RangeStyleResult GetStyle(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Applies one typed visual-formatting payload to one or more ranges.
    /// Validates all target addresses, protection, and known invalid settings before
    /// writing. Omitted properties preserve native state. Each selected border is
    /// independent; lineStyle none removes it. Theme colors follow workbook themes;
    /// fixed RGB remains fixed. A native failure does not promise rollback.
    /// </summary>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddresses">One or more target range addresses; all are validated before writing</param>
    /// <param name="formatOptions">Typed JSON object: fontName, fontSize, bold, italic, underline (none/single/double/singleAccounting/doubleAccounting), strikethrough, subscript, superscript, themeFont (0 none/1 major/2 minor), fontColor/fontThemeColor/fontTintAndShade, fillColor/fillThemeColor/fillTintAndShade, borders (position, lineStyle, weight, color/themeColor/tintAndShade), horizontalAlignment, verticalAlignment, wrapText, shrinkToFit, indentLevel (0-15), readingOrder (context/leftToRight/rightToLeft), orientation, numberFormat. Theme-color indices 1-12; tints -1 to 1. Border positions Left/Top/Bottom/Right/InsideHorizontal/InsideVertical/DiagonalUp/DiagonalDown. Nested keys remain camelCase.</param>
    [ServiceAction("format")]
    OperationResult Format(
        IExcelBatch batch,
        [AllowEmptyString] string sheetName,
        [RequiredParameter] string[] rangeAddresses,
        [RequiredParameter] CellFormatOptions formatOptions);

    // === VALIDATION OPERATIONS ===

    /// <summary>
    /// Adds data validation rules to range.
    /// Invalid type, comparison operator, and error style are rejected before replacing an existing rule.
    /// A later Excel error applying the replacement is not rolled back.
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
    OperationResult ValidateRange(
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
    OperationResult RemoveValidation(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === AUTO-FIT OPERATIONS ===

    /// <summary>
    /// Auto-fits column widths to content.
    /// Excel COM: Range.Columns.AutoFit()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Column range to auto-fit (e.g., 'A:D' or 'A1:D100')</param>
    [ServiceAction("auto-fit-columns")]
    OperationResult AutoFitColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Auto-fits row heights to content.
    /// Excel COM: Range.Rows.AutoFit()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row range to auto-fit (e.g., '1:10' or 'A1:D100')</param>
    [ServiceAction("auto-fit-rows")]
    OperationResult AutoFitRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === MERGE OPERATIONS ===

    /// <summary>
    /// Merges cells in range into a single cell.
    /// Excel COM: Range.Merge()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range to merge into a single cell (e.g., 'A1:D1')</param>
    [ServiceAction("merge-cells")]
    OperationResult MergeCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Unmerges previously merged cells.
    /// Excel COM: Range.UnMerge()
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range to unmerge (e.g., 'A1:D1')</param>
    [ServiceAction("unmerge-cells")]
    OperationResult UnmergeCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Checks if range contains merged cells and returns distinct merge areas.
    /// Excel COM: Range.MergeCells
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range to check for merged cells (e.g., 'A1:D10')</param>
    [ServiceAction("get-merge-info")]
    RangeMergeInfoResult GetMergeInfo(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === SIZING OPERATIONS ===

    /// <summary>
    /// Sets the width of columns in a range.
    /// Excel COM: Range.ColumnWidth property
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Column range to set width (e.g., 'A:A' or 'A1:D100')</param>
    /// <param name="columnWidth">Width in Excel character-width units, not points. Standard width is approximately 8.43. Range: 0.25-409.</param>
    [ServiceAction("set-column-width")]
    OperationResult SetColumnWidth(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] double columnWidth);

    /// <summary>
    /// Sets the height of rows in a range.
    /// Excel COM: Range.RowHeight property
    /// </summary>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row range to set height (e.g., '1:10' or 'A1:D100')</param>
    /// <param name="rowHeight">Height in points (1 point = 1/72 inch, approx 0.35mm). Default row height ~15 points. Range: 0-409 points.</param>
    [ServiceAction("set-row-height")]
    OperationResult SetRowHeight(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] double rowHeight);
}
