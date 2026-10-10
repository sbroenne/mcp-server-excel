using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Core range operations: get/set values and formulas, copy ranges, clear content, and discover data regions.
/// Use rangeedit for insert/delete/find/sort. Use rangeformat for styling/validation.
/// Use rangelink for hyperlinks and cell protection.
/// Calculation mode and explicit recalculation are handled by calculationmode.
/// Value/formula writes attempt to restore the prior calculation mode; manual mode needs explicit calculation.
/// Restoration can fail without failing the write; use get-settings when subsequent work depends on the mode.
/// Clear actions have no tool-level undo: clear-all removes values, formulas, and formats;
/// clear-contents removes values/formulas; clear-formats removes formats. Check the intended target.
///
/// Content writes and content-writing copies default to reject-nonempty: existing content stops before writing.
/// Use overwritePolicy='allow' only for intentional replacement. A failed inspection stops the write.
/// The check covers direct destinations, not future formula spills or transactional isolation.
/// Use 'clear-contents' (not 'clear-all') to preserve cell formatting when clearing data.
/// set-values preserves existing formatting; use set-number-format after if format change needed.
///
/// DATA FORMAT: values and formulas are 2D JSON arrays representing rows and columns.
/// Example: [[row1col1, row1col2], [row2col1, row2col2]]
/// Single cell returns [[value]] (always 2D).
/// Strict ISO dates such as "2025-01-15" are stored as native Excel dates.
/// Prefix the value with an apostrophe when an ISO-looking value must remain text.
///
/// REQUIRED PARAMETERS:
/// - sheetName + rangeAddress for cell operations (e.g., sheetName='Sheet1', rangeAddress='A1:D10')
/// - For named ranges, use sheetName='' (empty string) and rangeAddress='MyNamedRange'
///
/// COPY OPERATIONS: Specify source and target sheet/range for copy operations.
/// Copy requires an explicit pasteKind. Formats/validation preserve cell content.
/// Transpose is included in destination geometry; skipped blanks are not overwritten.
///
/// NUMBER FORMATS: Use US locale format codes (e.g., '#,##0.00', 'mm/dd/yyyy', '0.00%').
/// </summary>
[ServiceCategory("Range")]
[McpTool("range", Title = "Range Operations", Destructive = true, Category = "data",
    Description = "Write values and formulas, set number formats, copy ranges, and clear content or formatting. set-formulas accepts reference_style='a1' (default) or 'r1c1'; range addresses stay A1. Relative R1C1 references use each destination cell. " +
        "copy: Required paste_kind (all/values/formulas/formats/validation); transpose and skip_blanks default false. Formats/validation preserve content and need no overwrite permission. Formats include number formats, protection, and applicable conditional rules. All kinds require unmerged rectangular sources/destinations and a single-cell anchor or dimensions that are whole multiples of the source's paste dimensions, including transpose. Uses Excel's clipboard and clears owned copy mode on exit. " +
        "OVERWRITE POLICY: set-values, set-formulas, and content-writing copy kinds default to overwrite_policy='reject-nonempty'. Existing values, whitespace, errors, and formulas displaying blank are occupied. Copy checks exclude skipped source blanks. Conflicts or failed inspection stop before writing; errors list at most 10 conflicting addresses. Use overwrite_policy='allow' when the user's request authorizes replacement, without redundant confirmation. Never automatically retry a rejected write with allow. Checks cover direct destinations, including expanded copy targets, not future spills, rollback, or interactive Excel edits. " +
        "CLEAR ACTIONS HAVE NO TOOL-LEVEL UNDO: clear-all removes values, formulas, and formats; clear-contents removes values/formulas; clear-formats removes formats. Check the intended target before clearing. Use range_edit for insert/delete/find/sort/fill/auto-fill/create-series. Use range_format for styling/validation. Use range_link for hyperlinks/protection. " +
        "Value/formula writes attempt to restore the prior calculation mode; restoration can fail without failing the write. Verify the mode when subsequent work depends on it; manual mode needs explicit calculation. Use calculation_mode for recalculation. " +
        "EXCEL TABLES: If user asks to 'format as table', 'create a table', 'put data in an Excel Table' — do NOT try to use range for this. Use table(action:'create') on the data range to create a proper Excel Table with filter arrows, banded rows, and automatic expansion. " +
        "DATA FORMAT: 2D JSON arrays [[row1col1,row1col2],[row2col1,row2col2]]. Strict ISO dates such as '2025-01-15' are stored as native Excel dates; prefix an ISO-looking value with an apostrophe when it must remain text. In set-values, strings starting with '=' are written as formulas and every other cell keeps its value; prefix text with an apostrophe (\"'=\") when it must stay text. " +
        "MERGED CELLS: Writes that intersect merged cells fail unless the target is only the merged range's top-left cell; the error identifies affected merged ranges. " +
        "FILE INPUT: For set-values/set-formulas, provide EITHER inline values/formulas OR a valuesFile/formulasFile path to a .json or .csv file. Prefer file input for large datasets. Use clear-contents (not clear-all) to preserve formatting. NAMED RANGES: Use sheetName='' and rangeAddress=namedRangeName.")]
[McpReadOnlyActions("get-values", "get-formulas", "get-spill-info", "validate-formulas", "get-number-formats",
    "get-used-range", "get-current-region", "get-info", "get-special-cells", "trace-precedents", "trace-dependents")]
public interface IRangeCommands
{
    /// <summary>
    /// Traverses all native direct precedents reachable from every cell in the exact starting scope.
    /// Returns formulas/values, edges, cycles, and unresolved lookups, without a depth/output cap.
    /// Native getters return same-worksheet references only and do not fully resolve dynamic references.
    /// Ambiguous native absence is unresolved, not an empty successful lookup. Workbook coverage is
    /// never claimed complete. Does not activate/select cells, recalculate, or open external workbooks.
    /// coverage.workbookComplete is always false; a native no-range error is unresolved, not fabricated empty coverage.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name; empty for a named range</param>
    /// <param name="rangeAddress">Exact starting cells, rectangle, disjoint areas, or named range</param>
    [ServiceAction("trace-precedents")]
    RangeFormulaTraceResult TracePrecedents(IExcelBatch batch, [AllowEmptyString] string sheetName,
        [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Traverses all native direct dependents reachable from every cell in the exact starting scope.
    /// Returns formulas/values, edges, cycles, and unresolved lookups, without a depth/output cap.
    /// Coverage is native same-worksheet-only, never a complete workbook dependency graph.
    /// A missing native range is explicitly unresolved. No formula-text parsing, view changes,
    /// recalculation, or external workbook opening is performed.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name; empty for a named range</param>
    /// <param name="rangeAddress">Exact starting cells, rectangle, disjoint areas, or named range</param>
    [ServiceAction("trace-dependents")]
    RangeFormulaTraceResult TraceDependents(IExcelBatch batch, [AllowEmptyString] string sheetName,
        [RequiredParameter] string rangeAddress);

    // === VALUE OPERATIONS ===

    /// <summary>
    /// Gets values from a range as a 2D array. Formula errors are returned as canonical Excel
    /// names such as #REF! and are also listed in cellErrors with the affected cell, formula,
    /// raw COM code, and suggested fix.
    /// Single cell "A1" returns [[value]], range "A1:B2" returns [[v1,v2],[v3,v4]].
    /// Named ranges: Use empty sheetName and rangeAddress="NamedRange".
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name (e.g., 'SalesData')</param>
    [ServiceAction("get-values")]
    RangeValueResult GetValues(IExcelBatch batch, [AllowEmptyString] string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Sets values in a range from 2D array or file.
    /// Attempts to restore the prior calculation mode; restoration can fail without failing the write.
    /// Manual mode needs explicit calculation of dependent formulas.
    /// Provide EITHER values (inline JSON 2D array) OR valuesFile (path to .json or .csv file), not both.
    /// JSON file: must contain a 2D array like [[1,2],[3,4]].
    /// CSV file: rows become array rows, comma-separated values become columns.
    /// Every row must be rectangular and match the target range column count.
    /// Writes that intersect merged cells fail unless the target is only the merged range's top-left cell.
    /// Strict ISO dates such as "2025-01-15" are stored as native Excel dates.
    /// Prefix an ISO-looking value with an apostrophe to preserve it as text.
    /// Strings starting with "=" are written as formulas; all other cells in the same write keep their values.
    /// Prefix with an apostrophe ("'=") to store such a string as text.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only</param>
    /// <param name="rangeAddress">Cell range address matching data dimensions (e.g., 'A1' for [[value]], 'A1:B2' for [[v1,v2],[v3,v4]])</param>
    /// <param name="values">2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols). Strict ISO dates such as "2025-01-15" become native Excel dates. Optional if valuesFile is provided.</param>
    /// <param name="valuesFile">Path to a JSON or CSV file containing the values. JSON: 2D array. CSV: rows/columns. Alternative to inline values parameter.</param>
    /// <param name="overwritePolicy">reject-nonempty (default) checks all direct destinations and rejects existing content, including formulas displaying blank. allow permits intentional replacement, not bypassing Excel protection. Inspection failure stops the operation; no rollback or interactive-edit isolation.</param>
    [ServiceAction("set-values")]
    OperationResult SetValues(IExcelBatch batch, [AllowEmptyString] string sheetName, [RequiredParameter] string rangeAddress, List<List<object?>>? values = null, string? valuesFile = null, [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    // === FORMULA OPERATIONS ===

    /// <summary>
    /// Gets formulas from a range as a 2D array (empty string if no formula), together with
    /// calculated values. Formula errors are returned as canonical Excel names such as #REF!
    /// and are also listed in cellErrors with the affected cell, formula, raw COM code, and
    /// suggested fix.
    /// Single cell "A1" returns [["=SUM(B:B)"]], range "A1:B2" returns [[f1,f2],[f3,f4]].
    /// referenceStyle selects native A1 or R1C1 notation; rangeAddress remains A1.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name</param>
    /// <param name="referenceStyle">a1 (default) or r1c1 native formula notation; range addresses remain A1</param>
    [ServiceAction("get-formulas")]
    RangeFormulaResult GetFormulas(IExcelBatch batch, [AllowEmptyString] string sheetName,
        [RequiredParameter] string rangeAddress, [FromString] FormulaReferenceStyle referenceStyle = FormulaReferenceStyle.A1);

    /// <summary>
    /// Inspects every requested cell's native dynamic-array spill relationships.
    /// Returns ordinary/source/result/blocked states, source formulas, and established
    /// result extents without parsing formulas, changing selection, or recalculating.
    /// Blocked formulas have no invented extent. Sources may be outside the requested
    /// scope. Unsupported Excel sessions fail explicitly, not as empty ordinary cells.
    /// No preview limit; current relationships reflect Excel's current calculated state.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name, or empty string for a named range</param>
    /// <param name="rangeAddress">Exact range or named range to inspect completely</param>
    [ServiceAction("get-spill-info")]
    RangeSpillInfoResult GetSpillInfo(IExcelBatch batch, [AllowEmptyString] string sheetName,
        [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Sets formulas in a range from 2D array or file.
    /// Attempts to restore the prior calculation mode; restoration can fail without failing the write.
    /// Manual mode needs explicit calculation of dependent formulas.
    /// Provide EITHER formulas (inline JSON 2D array) OR formulasFile (path to .json file), not both.
    /// Every row must be rectangular and match the target range column count.
    /// Writes that intersect merged cells fail unless the target is only the merged range's top-left cell.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address matching formulas dimensions (e.g., 'A1:B2' for 2x2 formula array)</param>
    /// <param name="formulas">2D array of cells to set. Formulas need the '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']]). Cells may also be text, numbers, true/false, or null (empty cell), so labels and constants can sit beside formulas (e.g., [['Label', 5.86, true, null, '=1+1']]). Optional if formulasFile is provided.</param>
    /// <param name="formulasFile">Path to a JSON file containing the cells as a 2D array, with the same cell kinds as formulas. Alternative to inline formulas parameter.</param>
    /// <param name="overwritePolicy">reject-nonempty (default) rejects existing content before writing, including formulas displaying blank. allow permits authorized replacement. Checks cover direct destinations, not future formula spills; inspection failure stops the write.</param>
    /// <param name="referenceStyle">a1 (default) or r1c1 native formula notation; relative R1C1 references use each destination cell</param>
    [ServiceAction("set-formulas")]
    OperationResult SetFormulas(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress,
        List<List<object?>>? formulas = null, string? formulasFile = null,
        [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty,
        [FromString] FormulaReferenceStyle referenceStyle = FormulaReferenceStyle.A1);

    /// <summary>
    /// Validates formulas for syntax errors, undefined functions, and other issues without applying them.
    /// Detects common problems like undefined functions (e.g., GETVM3 without XA2. namespace),
    /// invalid references, syntax errors, and circular references.
    /// Accepts the same cells as set-formulas; text, numbers, true/false, and empty cells count as valid constants.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to validate</param>
    /// <param name="formulas">2D array of cells to validate. Formulas need the '=' prefix; text, numbers, true/false, and null are valid constants.</param>
    /// <param name="formulasFile">Path to a JSON file containing the cells to validate. Alternative to inline formulas parameter.</param>
    [ServiceAction("validate-formulas")]
    RangeFormulaValidationResult ValidateFormulas(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, List<List<object?>>? formulas = null, string? formulasFile = null);

    // === CLEAR OPERATIONS ===

    /// <summary>
    /// Clears all content (values, formulas, formats) from range.
    /// No tool-level undo. Check the intended target before clearing.
    /// Excel COM: Range.Clear()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to clear (e.g., 'A1:D10')</param>
    [ServiceAction("clear-all")]
    OperationResult ClearAll(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Clears only values and formulas (preserves formatting).
    /// No tool-level undo. Check the intended target before clearing.
    /// Excel COM: Range.ClearContents()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to clear (e.g., 'A1:D10')</param>
    [ServiceAction("clear-contents")]
    OperationResult ClearContents(IExcelBatch batch, [AllowEmptyString] string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Clears only formatting (preserves values and formulas).
    /// No tool-level undo. Check the intended target before clearing.
    /// Excel COM: Range.ClearFormats()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to clear (e.g., 'A1:D10')</param>
    [ServiceAction("clear-formats")]
    OperationResult ClearFormats(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === COPY OPERATIONS ===

    /// <summary>
    /// Copies a native rectangular range with an explicit paste kind.
    /// Formats/validation preserve content; formats also transfer number formats,
    /// protection, and applicable conditional rules. Formula paste includes constants
    /// and blanks. All kinds require unmerged rectangles and validated destination
    /// geometry, including transpose. Uses Excel's clipboard and clears the owned
    /// application's copy mode on exit. No rollback; saving remains explicit.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name for copy operations</param>
    /// <param name="sourceRange">Source range address for copy operations (e.g., 'A1:D10')</param>
    /// <param name="targetSheet">Target worksheet name for copy operations</param>
    /// <param name="targetRange">Target range address - can be single cell for paste destination (e.g., 'A1')</param>
    /// <param name="pasteKind">Required native paste kind: all, values, formulas, formats, or validation</param>
    /// <param name="transpose">Exchange source rows and columns; validation uses transposed destination dimensions</param>
    /// <param name="skipBlanks">Preserve destinations corresponding to native blank source cells; formulas displaying blank are not skipped</param>
    /// <param name="overwritePolicy">reject-nonempty (default) checks every content-writing destination, including expansion/repetition, but excludes skipped source blanks. allow permits intentional replacement. Formats/validation preserve content and need no overwrite permission. Neither policy bypasses sheet protection.</param>
    [ServiceAction("copy")]
    RangeCopyResult Copy(IExcelBatch batch, [RequiredParameter] string sourceSheet,
        [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet,
        [RequiredParameter] string targetRange, [RequiredParameter, FromString] PasteKind pasteKind,
        bool transpose = false, bool skipBlanks = false,
        [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    // === NUMBER FORMAT OPERATIONS ===

    /// <summary>
    /// Gets number format codes from range (2D array matching range dimensions).
    /// Excel COM: Range.NumberFormat
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    /// <returns>2D array of format codes (e.g., [["$#,##0.00", "0.00%"], ["m/d/yyyy", "General"]])</returns>
    [ServiceAction("get-number-formats")]
    RangeNumberFormatResult GetNumberFormats(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Sets uniform number format for entire range.
    /// Excel COM: Range.NumberFormat = formatCode
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    /// <param name="formatCode">Number format code in US locale (e.g., '#,##0.00' for numbers, 'mm/dd/yyyy' for dates, '0.00%' for percentages, 'General' for default, '@' for text)</param>
    [ServiceAction("set-number-format")]
    OperationResult SetNumberFormat(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string formatCode);

    /// <summary>
    /// Sets number formats cell-by-cell from 2D array or file.
    /// Provide EITHER formats (inline JSON 2D array) OR formatsFile (path to .json file), not both.
    /// Excel COM: Range.NumberFormat (per cell)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address matching formats dimensions</param>
    /// <param name="formats">2D array of format codes - same dimensions as target range (e.g., [['#,##0.00', '0.00%'], ['mm/dd/yyyy', 'General']]). Optional if formatsFile is provided.</param>
    /// <param name="formatsFile">Path to a JSON file containing 2D array of format codes. Alternative to inline formats parameter.</param>
    [ServiceAction("set-number-formats")]
    OperationResult SetNumberFormats(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, List<List<string>>? formats = null, string? formatsFile = null);

    // === DISCOVERY OPERATIONS ===

    /// <summary>
    /// Gets the used range (all non-empty cells) from worksheet.
    /// Excel COM: Worksheet.UsedRange
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [ServiceAction("get-used-range")]
    RangeValueResult GetUsedRange(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Gets the current region (contiguous data block) around a cell.
    /// Excel COM: Range.CurrentRegion
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Single cell address (e.g., 'B5') - expands to contiguous data region around this cell</param>
    [ServiceAction("get-current-region")]
    RangeValueResult GetCurrentRegion(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress);

    /// <summary>
    /// Gets range information (address, dimensions, number formats).
    /// Excel COM: Range.Address, Range.Rows.Count, Range.Columns.Count, Range.NumberFormat
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    [ServiceAction("get-info")]
    RangeInfoResult GetInfo(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Finds all formulas, constants, truly blank cells, error cells, or visible cells
    /// in exactly the requested range. Returns complete matching area addresses and
    /// their total cell count without a preview limit. Formula results displaying
    /// empty text are formulas, not blanks. Errors include constants and formulas.
    /// Visible excludes cells in hidden or filtered rows and hidden columns.
    /// A single-cell request never expands to the sheet's used range.
    /// No matches returns an empty successful result; invalid inputs and Excel
    /// failures remain errors. Does not select cells or change workbook content.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name, or empty string for a named range</param>
    /// <param name="rangeAddress">Exact range or named range to inspect; all matching areas are returned</param>
    /// <param name="cellKind">Cell selector: formulas, constants, blanks, errors, or visible</param>
    [ServiceAction("get-special-cells")]
    SpecialCellsResult GetSpecialCells(IExcelBatch batch, [AllowEmptyString] string sheetName,
        [RequiredParameter] string rangeAddress, [RequiredParameter, FromString] SpecialCellKind cellKind);
}

// === SUPPORTING TYPES (shared by all range interfaces) ===

/// <summary>
/// Direction to shift cells when inserting
/// </summary>
public enum InsertShiftDirection
{
    /// <summary>Shift existing cells down</summary>
    Down,
    /// <summary>Shift existing cells right</summary>
    Right
}

/// <summary>
/// Direction to shift cells when deleting
/// </summary>
public enum DeleteShiftDirection
{
    /// <summary>Shift remaining cells up</summary>
    Up,
    /// <summary>Shift remaining cells left</summary>
    Left
}

/// <summary>
/// Options for find operations
/// </summary>
public class FindOptions
{
    /// <summary>Whether to match case</summary>
    public bool MatchCase { get; set; }

    /// <summary>Whether to match entire cell content</summary>
    public bool MatchEntireCell { get; set; }

    /// <summary>Whether to search in formulas</summary>
    public bool SearchFormulas { get; set; } = true;

    /// <summary>Whether to search in values</summary>
    public bool SearchValues { get; set; } = true;

    /// <summary>Whether to search in comments</summary>
    public bool SearchComments { get; set; }
}

/// <summary>
/// Options for replace operations
/// </summary>
public class ReplaceOptions : FindOptions
{
    /// <summary>Whether to replace all occurrences (true) or just first (false)</summary>
    public bool ReplaceAll { get; set; } = true;
}

/// <summary>
/// Sort column definition
/// </summary>
public class SortColumn
{
    /// <summary>Column index within range (1-based)</summary>
    public int ColumnIndex { get; set; }

    /// <summary>Sort direction (true = ascending, false = descending)</summary>
    public bool Ascending { get; set; } = true;
}
