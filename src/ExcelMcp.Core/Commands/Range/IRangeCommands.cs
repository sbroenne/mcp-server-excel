using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Core range operations: get/set values and formulas, copy ranges, clear content, and discover data regions.
/// Use rangeedit for insert/delete/find/sort. Use rangeformat for styling/validation.
/// Use rangelink for hyperlinks and cell protection.
/// Calculation mode and explicit recalculation are handled by calculationmode.
///
/// BEST PRACTICE: Use 'get-values' to check existing data before overwriting.
/// Use 'clear-contents' (not 'clear-all') to preserve cell formatting when clearing data.
/// set-values preserves existing formatting; use set-number-format after if format change needed.
///
/// DATA FORMAT: values and formulas are 2D JSON arrays representing rows and columns.
/// Example: [[row1col1, row1col2], [row2col1, row2col2]]
/// Single cell returns [[value]] (always 2D).
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
[McpTool("range", Title = "Range Operations", Destructive = true, Category = "data",
    Description = "Core range operations: get/set values and formulas, copy ranges, clear content, discover data regions. Use range_edit for insert/delete/find/sort. Use range_format for styling/validation. Use range_link for hyperlinks/protection. Use calculation_mode for recalculation. EXCEL TABLES: If user asks to 'format as table', 'create a table', 'put data in an Excel Table' — do NOT try to use range for this. Use table(action:'create') on the data range to create a proper Excel Table with filter arrows, banded rows, and automatic expansion. DATA FORMAT: 2D JSON arrays [[row1col1,row1col2],[row2col1,row2col2]]. Single cell returns [[value]]. FILE INPUT: For set-values/set-formulas, provide EITHER inline values/formulas OR a valuesFile/formulasFile path to a .json or .csv file. Prefer file input for large datasets. BEST PRACTICE: get-values before overwriting, clear-contents (not clear-all) to preserve formatting. NAMED RANGES: Use sheetName='' and rangeAddress=namedRangeName.")]
public interface IRangeCommands
{
    // === VALUE OPERATIONS ===

    /// <summary>
    /// Gets values from a range as 2D array.
    /// Single cell "A1" returns [[value]], range "A1:B2" returns [[v1,v2],[v3,v4]].
    /// Named ranges: Use empty sheetName and rangeAddress="NamedRange".
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name (e.g., 'SalesData')</param>
    [ServiceAction("get-values")]
    RangeValueResult GetValues(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Sets values in a range from 2D array or file.
    /// Provide EITHER values (inline JSON 2D array) OR valuesFile (path to .json or .csv file), not both.
    /// JSON file: must contain a 2D array like [[1,2],[3,4]].
    /// CSV file: rows become array rows, comma-separated values become columns.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range - REQUIRED for cell addresses, use empty string for named ranges only</param>
    /// <param name="rangeAddress">Cell range address matching data dimensions (e.g., 'A1' for [[value]], 'A1:B2' for [[v1,v2],[v3,v4]])</param>
    /// <param name="values">2D array of values to set - rows are outer array, columns are inner array (e.g., [[1,2,3],[4,5,6]] for 2 rows x 3 cols). Optional if valuesFile is provided.</param>
    /// <param name="valuesFile">Path to a JSON or CSV file containing the values. JSON: 2D array. CSV: rows/columns. Alternative to inline values parameter.</param>
    [ServiceAction("set-values")]
    OperationResult SetValues(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, List<List<object?>>? values = null, string? valuesFile = null);

    // === FORMULA OPERATIONS ===

    /// <summary>
    /// Gets formulas from a range as 2D array (empty string if no formula).
    /// Single cell "A1" returns [["=SUM(B:B)"]], range "A1:B2" returns [[f1,f2],[f3,f4]].
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address (e.g., 'A1', 'A1:D10', 'B:D') or named range name</param>
    [ServiceAction("get-formulas")]
    RangeFormulaResult GetFormulas(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Sets formulas in a range from 2D array or file.
    /// Provide EITHER formulas (inline JSON 2D array) OR formulasFile (path to .json file), not both.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address matching formulas dimensions (e.g., 'A1:B2' for 2x2 formula array)</param>
    /// <param name="formulas">2D array of formulas to set - include '=' prefix (e.g., [['=A1+B1', '=SUM(A:A)'], ['=C1*2', '=AVERAGE(B:B)']]). Optional if formulasFile is provided.</param>
    /// <param name="formulasFile">Path to a JSON file containing the formulas as a 2D array. Alternative to inline formulas parameter.</param>
    [ServiceAction("set-formulas")]
    OperationResult SetFormulas(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, List<List<string>>? formulas = null, string? formulasFile = null);

    // === CLEAR OPERATIONS ===

    /// <summary>
    /// Clears all content (values, formulas, formats) from range.
    /// Excel COM: Range.Clear()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to clear (e.g., 'A1:D10')</param>
    [ServiceAction("clear-all")]
    OperationResult ClearAll(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Clears only values and formulas (preserves formatting).
    /// Excel COM: Range.ClearContents()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to clear (e.g., 'A1:D10')</param>
    [ServiceAction("clear-contents")]
    OperationResult ClearContents(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Clears only formatting (preserves values and formulas).
    /// Excel COM: Range.ClearFormats()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to clear (e.g., 'A1:D10')</param>
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
    [ServiceAction("copy")]
    OperationResult Copy(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange);

    /// <summary>
    /// Copies only values (no formulas or formatting).
    /// Excel COM: Range.PasteSpecial(xlPasteValues)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name for copy operations</param>
    /// <param name="sourceRange">Source range address for copy operations (e.g., 'A1:D10')</param>
    /// <param name="targetSheet">Target worksheet name for copy operations</param>
    /// <param name="targetRange">Target range address - can be single cell for paste destination (e.g., 'A1')</param>
    [ServiceAction("copy-values")]
    OperationResult CopyValues(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange);

    /// <summary>
    /// Copies only formulas (no values or formatting).
    /// Excel COM: Range.PasteSpecial(xlPasteFormulas)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sourceSheet">Source worksheet name for copy operations</param>
    /// <param name="sourceRange">Source range address for copy operations (e.g., 'A1:D10')</param>
    /// <param name="targetSheet">Target worksheet name for copy operations</param>
    /// <param name="targetRange">Target range address - can be single cell for paste destination (e.g., 'A1')</param>
    [ServiceAction("copy-formulas")]
    OperationResult CopyFormulas(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange);

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



