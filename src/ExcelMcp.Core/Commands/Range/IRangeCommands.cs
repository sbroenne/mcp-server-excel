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
/// Restoration can fail without failing the write; use get-mode when subsequent work depends on the mode.
/// Clear actions have no tool-level undo: clear-all removes values, formulas, and formats;
/// clear-contents removes values/formulas; clear-formats removes formats. Check the intended target.
///
/// Content writes/copies default to reject-nonempty: existing content stops the operation before writing.
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
///
/// NUMBER FORMATS: Use US locale format codes (e.g., '#,##0.00', 'mm/dd/yyyy', '0.00%').
/// </summary>
[ServiceCategory("range", "Range")]
[MacCapability(MacCapabilityTier.Unsupported, MacImplementationStatus.Blocked, false,
    Evidence = "The interface default covers range actions without a separately verified native route. Formula validation requires Excel's parser and localized error behavior, which the reviewed macOS surfaces do not expose as a non-mutating contract.",
    ExcelApiVersion = "Excel for Mac 16.113.1 Apple Events dictionary; Office.js ExcelApi through 1.21.",
    Blocker = "current supported macOS APIs cannot preserve this range contract; use the Windows COM backend")]
[McpTool("range", Title = "Range Operations", Destructive = true, Category = "data",
    Description = "Core range operations: get/set values and formulas, copy ranges, clear content, discover data regions. OVERWRITE POLICY: set-values, set-formulas, copy, copy-values, and copy-formulas default to overwrite_policy='reject-nonempty'. Existing values, whitespace, errors, and formulas displaying blank are occupied. Conflicts or failed inspection stop before writing; errors list at most 10 conflicting addresses. Use overwrite_policy='allow' when the user's request authorizes replacement, without redundant confirmation. Never automatically retry a rejected write with allow. Checks cover direct destinations, including expanded copy targets, not future spills, rollback, or interactive Excel edits. Protected copies require unmerged rectangular sources/destinations and a single-cell anchor or destination dimensions that are whole multiples of the source. CLEAR ACTIONS HAVE NO TOOL-LEVEL UNDO: clear-all removes values, formulas, and formats; clear-contents removes values/formulas; clear-formats removes formats. Check the intended target before clearing. Use range_edit for insert/delete/find/sort. Use range_format for styling/validation. Use range_link for hyperlinks/protection. Value/formula writes attempt to restore the prior calculation mode; restoration can fail without failing the write. Use calculation_mode get-mode when subsequent work depends on the mode; manual mode needs explicit calculation. Use calculation_mode for recalculation. EXCEL TABLES: If user asks to 'format as table', 'create a table', 'put data in an Excel Table' — do NOT try to use range for this. Use table(action:'create') on the data range to create a proper Excel Table with filter arrows, banded rows, and automatic expansion. DATA FORMAT: 2D JSON arrays [[row1col1,row1col2],[row2col1,row2col2]]. Single cell returns [[value]]. Strict ISO dates such as '2025-01-15' are stored as native Excel dates; prefix an ISO-looking value with an apostrophe when it must remain text. MERGED CELLS: Writes that intersect merged cells fail unless the target is only the merged range's top-left cell; the error identifies affected merged ranges. FILE INPUT: For set-values/set-formulas, provide EITHER inline values/formulas OR a valuesFile/formulasFile path to a .json or .csv file. Prefer file input for large datasets. Use clear-contents (not clear-all) to preserve formatting. NAMED RANGES: Use sheetName='' and rangeAddress=namedRangeName.")]
public interface IRangeCommands
{
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
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
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
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only</param>
    /// <param name="rangeAddress">Cell range address matching data dimensions (e.g., 'A1' for [[value]], 'A1:B2' for [[v1,v2],[v3,v4]])</param>
    /// <param name="values">2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols). Strict ISO dates such as "2025-01-15" become native Excel dates. Optional if valuesFile is provided.</param>
    /// <param name="valuesFile">Path to a JSON or CSV file containing the values. JSON: 2D array. CSV: rows/columns. Alternative to inline values parameter.</param>
    /// <param name="overwritePolicy">reject-nonempty (default) checks all direct destinations and rejects existing content, including formulas displaying blank. allow permits intentional replacement, not bypassing Excel protection. Inspection failure stops the operation; no rollback or interactive-edit isolation.</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [ServiceAction("set-values")]
    OperationResult SetValues(IExcelBatch batch, [AllowEmptyString] string sheetName, [RequiredParameter] string rangeAddress, List<List<object?>>? values = null, string? valuesFile = null, [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    // === FORMULA OPERATIONS ===

    /// <summary>
    /// Gets formulas from a range as a 2D array (empty string if no formula), together with
    /// calculated values. Formula errors are returned as canonical Excel names such as #REF!
    /// and are also listed in cellErrors with the affected cell, formula, raw COM code, and
    /// suggested fix.
    /// Single cell "A1" returns [["=SUM(B:B)"]], range "A1:B2" returns [[f1,f2],[f3,f4]].
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [ServiceAction("get-formulas")]
    RangeFormulaResult GetFormulas(IExcelBatch batch, [AllowEmptyString] string sheetName, [RequiredParameter] string rangeAddress);

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
    /// <param name="formulas">2D array of formulas to set - include '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']]). Optional if formulasFile is provided.</param>
    /// <param name="formulasFile">Path to a JSON file containing the formulas as a 2D array. Alternative to inline formulas parameter.</param>
    /// <param name="overwritePolicy">reject-nonempty (default) rejects existing content before writing, including formulas displaying blank. allow permits authorized replacement. Checks cover direct destinations, not future formula spills; inspection failure stops the write.</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [ServiceAction("set-formulas")]
    OperationResult SetFormulas(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, List<List<string>>? formulas = null, string? formulasFile = null, [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    /// <summary>
    /// Validates formulas for syntax errors, undefined functions, and other issues without applying them.
    /// Detects common problems like undefined functions (e.g., GETVM3 without XA2. namespace),
    /// invalid references, syntax errors, and circular references.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to validate</param>
    /// <param name="formulas">2D array of formulas to validate - include '=' prefix</param>
    /// <param name="formulasFile">Path to a JSON file containing the formulas to validate. Alternative to inline formulas parameter.</param>
    [ServiceAction("validate-formulas")]
    RangeFormulaValidationResult ValidateFormulas(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, List<List<string>>? formulas = null, string? formulasFile = null);

    // === CLEAR OPERATIONS ===

    /// <summary>
    /// Clears all content (values, formulas, formats) from range.
    /// No tool-level undo. Check the intended target before clearing.
    /// Excel COM: Range.Clear()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to clear (e.g., 'A1:D10')</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
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
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
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
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
    [ServiceAction("clear-formats")]
    OperationResult ClearFormats(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === COPY OPERATIONS ===

    /// <summary>
    /// Copies range to another location (all content).
    /// Excel COM: Range.Copy()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name for copy operations</param>
    /// <param name="sourceRange">Source range address for copy operations (e.g., 'A1:D10')</param>
    /// <param name="targetSheet">Target worksheet name for copy operations</param>
    /// <param name="targetRange">Target range address - can be single cell for paste destination (e.g., 'A1')</param>
    /// <param name="overwritePolicy">reject-nonempty (default) checks the entire paste destination, including expansion/repetition and cells cleared by source blanks. Protected copies require unmerged rectangles and compatible dimensions. allow permits intentional replacement; inspection errors stop protected writes.</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "CLI and MCP copied complete range content through the declared destination range command.",
        ExcelApiVersion = "Excel for Mac 16.113.1.")]
    [ServiceAction("copy")]
    OperationResult Copy(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange, [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    /// <summary>
    /// Copies only values (no formulas or formatting).
    /// Excel COM: Range.PasteSpecial(xlPasteValues)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name for copy operations</param>
    /// <param name="sourceRange">Source range address for copy operations (e.g., 'A1:D10')</param>
    /// <param name="targetSheet">Target worksheet name for copy operations</param>
    /// <param name="targetRange">Target range address - can be single cell for paste destination (e.g., 'A1')</param>
    /// <param name="overwritePolicy">reject-nonempty (default) checks the entire paste destination, including expansion/repetition and cells cleared by source blanks. Protected copies require unmerged rectangles and compatible dimensions. allow permits intentional replacement.</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "CLI and MCP copied values without formulas or formatting through direct value matrices.",
        ExcelApiVersion = "Excel for Mac 16.113.1.")]
    [ServiceAction("copy-values")]
    OperationResult CopyValues(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange, [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    /// <summary>
    /// Copies formulas and source constants (no formatting), following Excel formula-paste semantics.
    /// Excel COM: Range.PasteSpecial(xlPasteFormulas)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name for copy operations</param>
    /// <param name="sourceRange">Source range address for copy operations (e.g., 'A1:D10')</param>
    /// <param name="targetSheet">Target worksheet name for copy operations</param>
    /// <param name="targetRange">Target range address - can be single cell for paste destination (e.g., 'A1')</param>
    /// <param name="overwritePolicy">reject-nonempty (default) checks the entire paste destination, including expansion/repetition and cells cleared by source blanks. Excel formula paste also copies source constants. Protected copies require unmerged rectangles and compatible dimensions. allow permits intentional replacement.</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "CLI and MCP copied formulas without formatting through R1C1 matrices and preserved relative-reference adjustment.",
        ExcelApiVersion = "Excel for Mac 16.113.1.")]
    [ServiceAction("copy-formulas")]
    OperationResult CopyFormulas(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange, [FromString] OverwritePolicy overwritePolicy = OverwritePolicy.RejectNonempty);

    // === NUMBER FORMAT OPERATIONS ===

    /// <summary>
    /// Gets number format codes from range (2D array matching range dimensions).
    /// Excel COM: Range.NumberFormat
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    /// <returns>2D array of format codes (e.g., [["$#,##0.00", "0.00%"], ["m/d/yyyy", "General"]])</returns>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
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
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true)]
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
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "CLI and MCP applied and independently read back mixed two-dimensional number-format matrices.",
        ExcelApiVersion = "Excel for Mac 16.113.1.")]
    [ServiceAction("set-number-formats")]
    OperationResult SetNumberFormats(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, List<List<string>>? formats = null, string? formatsFile = null);

    // === DISCOVERY OPERATIONS ===

    /// <summary>
    /// Gets the used range (all non-empty cells) from worksheet.
    /// Excel COM: Worksheet.UsedRange
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Blocked, false,
        Evidence = "On populated sheets, CLI and MCP both received the empty-sheet $A$1 fallback because JXA used range returned a missing object. JXA special cells also returned a missing object and typed AppleScript returned parameter error -50.",
        ExcelApiVersion = "Excel for Mac 16.113.1.",
        Blocker = "the native routes cannot return a populated live used range without approximating Worksheet.UsedRange semantics")]
    [ServiceAction("get-used-range")]
    RangeValueResult GetUsedRange(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Gets the current region (contiguous data block) around a cell.
    /// Excel COM: Range.CurrentRegion
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Single cell address (e.g., 'B5') - expands to contiguous data region around this cell</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Blocked, false,
        Evidence = "Native CurrentRegion probes returned a missing object through JXA and parameter error -50 through typed AppleScript.",
        ExcelApiVersion = "Excel for Mac 16.113.1.",
        Blocker = "the native current-region candidate has no verified route and has not completed real CLI and MCP acceptance")]
    [ServiceAction("get-current-region")]
    RangeValueResult GetCurrentRegion(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress);

    /// <summary>
    /// Gets range information (address, dimensions, number formats).
    /// Excel COM: Range.Address, Range.Rows.Count, Range.Columns.Count, Range.NumberFormat
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1:D10')</param>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "CLI and MCP returned absolute address, dimensions, number format, and positive range geometry.",
        ExcelApiVersion = "Excel for Mac 16.113.1.")]
    [ServiceAction("get-info")]
    RangeInfoResult GetInfo(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);
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
