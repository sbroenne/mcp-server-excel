using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Excel range operation commands - unified API for all range data operations
/// Single cell is treated as 1x1 range. Named ranges work transparently via rangeAddress parameter.
/// All operations are COM-backed (no data processing in server).
/// </summary>
public interface IRangeCommands
{
    // === VALUE OPERATIONS ===

    /// <summary>
    /// Gets values from a range as 2D array
    /// Single cell "A1" returns [[value]], range "A1:B2" returns [[v1,v2],[v3,v4]]
    /// Named ranges: Use empty sheetName and rangeAddress="NamedRange"
    /// </summary>
    Task<RangeValueResult> GetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Sets values in a range from 2D array
    /// Single cell "A1" accepts [[value]], range "A1:B2" accepts [[v1,v2],[v3,v4]]
    /// </summary>
    Task<OperationResult> SetValuesAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<object?>> values);

    // === FORMULA OPERATIONS ===

    /// <summary>
    /// Gets formulas from a range as 2D array (empty string if no formula)
    /// Single cell "A1" returns [["=SUM(B:B)"]], range "A1:B2" returns [[f1,f2],[f3,f4]]
    /// </summary>
    Task<RangeFormulaResult> GetFormulasAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Sets formulas in a range from 2D array
    /// </summary>
    Task<OperationResult> SetFormulasAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formulas);

    // === CLEAR OPERATIONS ===

    /// <summary>
    /// Clears all content (values, formulas, formats) from range
    /// Excel COM: Range.Clear()
    /// </summary>
    Task<OperationResult> ClearAllAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Clears only values and formulas (preserves formatting)
    /// Excel COM: Range.ClearContents()
    /// </summary>
    Task<OperationResult> ClearContentsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Clears only formatting (preserves values and formulas)
    /// Excel COM: Range.ClearFormats()
    /// </summary>
    Task<OperationResult> ClearFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    // === COPY OPERATIONS ===

    /// <summary>
    /// Copies range to another location (all content)
    /// Excel COM: Range.Copy()
    /// </summary>
    Task<OperationResult> CopyAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange);

    /// <summary>
    /// Copies only values (no formulas or formatting)
    /// Excel COM: Range.PasteSpecial(xlPasteValues)
    /// </summary>
    Task<OperationResult> CopyValuesAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange);

    /// <summary>
    /// Copies only formulas (no values or formatting)
    /// Excel COM: Range.PasteSpecial(xlPasteFormulas)
    /// </summary>
    Task<OperationResult> CopyFormulasAsync(IExcelBatch batch, string sourceSheet, string sourceRange, string targetSheet, string targetRange);

    // === INSERT/DELETE OPERATIONS ===

    /// <summary>
    /// Inserts blank cells, shifting existing cells down or right
    /// Excel COM: Range.Insert(shift)
    /// </summary>
    Task<OperationResult> InsertCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress, InsertShiftDirection shift);

    /// <summary>
    /// Deletes cells, shifting remaining cells up or left
    /// Excel COM: Range.Delete(shift)
    /// </summary>
    Task<OperationResult> DeleteCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress, DeleteShiftDirection shift);

    /// <summary>
    /// Inserts entire rows above the range
    /// Excel COM: Range.EntireRow.Insert()
    /// </summary>
    Task<OperationResult> InsertRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Deletes entire rows in the range
    /// Excel COM: Range.EntireRow.Delete()
    /// </summary>
    Task<OperationResult> DeleteRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Inserts entire columns to the left of the range
    /// Excel COM: Range.EntireColumn.Insert()
    /// </summary>
    Task<OperationResult> InsertColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Deletes entire columns in the range
    /// Excel COM: Range.EntireColumn.Delete()
    /// </summary>
    Task<OperationResult> DeleteColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    // === FIND/REPLACE OPERATIONS ===

    /// <summary>
    /// Finds all cells matching criteria in range
    /// Excel COM: Range.Find()
    /// </summary>
    Task<RangeFindResult> FindAsync(IExcelBatch batch, string sheetName, string rangeAddress, string searchValue, FindOptions options);

    /// <summary>
    /// Replaces text/values in range
    /// Excel COM: Range.Replace()
    /// </summary>
    Task<OperationResult> ReplaceAsync(IExcelBatch batch, string sheetName, string rangeAddress, string findValue, string replaceValue, ReplaceOptions options);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts range by one or more columns
    /// Excel COM: Range.Sort()
    /// </summary>
    Task<OperationResult> SortAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<SortColumn> sortColumns, bool hasHeaders = true);

    // === NATIVE EXCEL COM OPERATIONS (AI/LLM ESSENTIAL) ===

    /// <summary>
    /// Gets the used range (all non-empty cells) from worksheet
    /// Excel COM: Worksheet.UsedRange
    /// </summary>
    Task<RangeValueResult> GetUsedRangeAsync(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Gets the current region (contiguous data block) around a cell
    /// Excel COM: Range.CurrentRegion
    /// </summary>
    Task<RangeValueResult> GetCurrentRegionAsync(IExcelBatch batch, string sheetName, string cellAddress);

    /// <summary>
    /// Gets range information (address, dimensions, number formats)
    /// Excel COM: Range.Address, Range.Rows.Count, Range.Columns.Count, Range.NumberFormat
    /// </summary>
    Task<RangeInfoResult> GetRangeInfoAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    // === HYPERLINK OPERATIONS ===

    /// <summary>
    /// Adds hyperlink to a single cell
    /// Excel COM: Worksheet.Hyperlinks.Add()
    /// </summary>
    Task<OperationResult> AddHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress, string url, string? displayText = null, string? tooltip = null);

    /// <summary>
    /// Removes hyperlink from a single cell or all hyperlinks from a range
    /// Excel COM: Range.Hyperlinks.Delete()
    /// </summary>
    Task<OperationResult> RemoveHyperlinkAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Lists all hyperlinks in a worksheet
    /// Excel COM: Worksheet.Hyperlinks collection
    /// </summary>
    Task<RangeHyperlinkResult> ListHyperlinksAsync(IExcelBatch batch, string sheetName);

    /// <summary>
    /// Gets hyperlink from a specific cell
    /// Excel COM: Range.Hyperlink
    /// </summary>
    Task<RangeHyperlinkResult> GetHyperlinkAsync(IExcelBatch batch, string sheetName, string cellAddress);
}

// === SUPPORTING TYPES ===

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
    public bool MatchCase { get; set; } = false;

    /// <summary>Whether to match entire cell content</summary>
    public bool MatchEntireCell { get; set; } = false;

    /// <summary>Whether to search in formulas</summary>
    public bool SearchFormulas { get; set; } = true;

    /// <summary>Whether to search in values</summary>
    public bool SearchValues { get; set; } = true;

    /// <summary>Whether to search in comments</summary>
    public bool SearchComments { get; set; } = false;
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
