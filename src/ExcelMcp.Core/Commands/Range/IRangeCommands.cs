using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Core Excel range data operations - values, formulas, copy, clear, discovery.
/// Single cell is treated as 1x1 range. Named ranges work transparently via rangeAddress parameter.
/// All operations are COM-backed (no data processing in server).
/// Use IRangeEditCommands for insert/delete/find/sort.
/// Use IRangeFormatCommands for styling/validation/merge.
/// Use IRangeLinkCommands for hyperlinks and cell protection.
/// </summary>
[ServiceCategory("range", "Range")]
[McpTool("excel_range")]
public interface IRangeCommands
{
    // === VALUE OPERATIONS ===

    /// <summary>
    /// Gets values from a range as 2D array.
    /// Single cell "A1" returns [[value]], range "A1:B2" returns [[v1,v2],[v3,v4]].
    /// Named ranges: Use empty sheetName and rangeAddress="NamedRange".
    /// </summary>
    [ServiceAction("get-values")]
    RangeValueResult GetValues(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Sets values in a range from 2D array.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Target worksheet name</param>
    /// <param name="rangeAddress">
    /// MUST specify full range matching data dimensions:
    /// - Single cell: "A1" for [[value]]
    /// - Multi-cell: "A1:B2" for [[v1,v2],[v3,v4]]
    /// IMPORTANT: Always specify the exact range address.
    /// </param>
    /// <param name="values">2D array of values to set</param>
    [ServiceAction("set-values")]
    OperationResult SetValues(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] List<List<object?>> values);

    // === FORMULA OPERATIONS ===

    /// <summary>
    /// Gets formulas from a range as 2D array (empty string if no formula).
    /// Single cell "A1" returns [["=SUM(B:B)"]], range "A1:B2" returns [[f1,f2],[f3,f4]].
    /// </summary>
    [ServiceAction("get-formulas")]
    RangeFormulaResult GetFormulas(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Sets formulas in a range from 2D array.
    /// </summary>
    [ServiceAction("set-formulas")]
    OperationResult SetFormulas(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] List<List<string>> formulas);

    // === CLEAR OPERATIONS ===

    /// <summary>
    /// Clears all content (values, formulas, formats) from range.
    /// Excel COM: Range.Clear()
    /// </summary>
    [ServiceAction("clear-all")]
    OperationResult ClearAll(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Clears only values and formulas (preserves formatting).
    /// Excel COM: Range.ClearContents()
    /// </summary>
    [ServiceAction("clear-contents")]
    OperationResult ClearContents(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Clears only formatting (preserves values and formulas).
    /// Excel COM: Range.ClearFormats()
    /// </summary>
    [ServiceAction("clear-formats")]
    OperationResult ClearFormats(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === COPY OPERATIONS ===

    /// <summary>
    /// Copies range to another location (all content).
    /// Excel COM: Range.Copy()
    /// </summary>
    [ServiceAction("copy")]
    OperationResult Copy(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange);

    /// <summary>
    /// Copies only values (no formulas or formatting).
    /// Excel COM: Range.PasteSpecial(xlPasteValues)
    /// </summary>
    [ServiceAction("copy-values")]
    OperationResult CopyValues(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange);

    /// <summary>
    /// Copies only formulas (no values or formatting).
    /// Excel COM: Range.PasteSpecial(xlPasteFormulas)
    /// </summary>
    [ServiceAction("copy-formulas")]
    OperationResult CopyFormulas(IExcelBatch batch, [RequiredParameter] string sourceSheet, [RequiredParameter] string sourceRange, [RequiredParameter] string targetSheet, [RequiredParameter] string targetRange);

    // === NUMBER FORMAT OPERATIONS ===

    /// <summary>
    /// Gets number format codes from range (2D array matching range dimensions).
    /// Excel COM: Range.NumberFormat
    /// </summary>
    /// <returns>2D array of format codes (e.g., [["$#,##0.00", "0.00%"], ["m/d/yyyy", "General"]])</returns>
    [ServiceAction("get-number-formats")]
    RangeNumberFormatResult GetNumberFormats(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Sets uniform number format for entire range.
    /// Excel COM: Range.NumberFormat = formatCode
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <param name="formatCode">Excel format code (e.g., "$#,##0.00", "0.00%", "m/d/yyyy", "General", "@")</param>
    [ServiceAction("set-number-format")]
    OperationResult SetNumberFormat(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string formatCode);

    /// <summary>
    /// Sets number formats cell-by-cell from 2D array.
    /// Excel COM: Range.NumberFormat (per cell)
    /// </summary>
    [ServiceAction("set-number-formats")]
    OperationResult SetNumberFormats(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] List<List<string>> formats);

    // === DISCOVERY OPERATIONS ===

    /// <summary>
    /// Gets the used range (all non-empty cells) from worksheet.
    /// Excel COM: Worksheet.UsedRange
    /// </summary>
    [ServiceAction("get-used-range")]
    RangeValueResult GetUsedRange(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Gets the current region (contiguous data block) around a cell.
    /// Excel COM: Range.CurrentRegion
    /// </summary>
    [ServiceAction("get-current-region")]
    RangeValueResult GetCurrentRegion(IExcelBatch batch, string sheetName, [RequiredParameter] string cellAddress);

    /// <summary>
    /// Gets range information (address, dimensions, number formats).
    /// Excel COM: Range.Address, Range.Rows.Count, Range.Columns.Count, Range.NumberFormat
    /// </summary>
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



