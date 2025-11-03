using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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

    // === NUMBER FORMAT OPERATIONS ===

    /// <summary>
    /// Gets number format codes from range (2D array matching range dimensions)
    /// Excel COM: Range.NumberFormat
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <returns>2D array of format codes (e.g., [["$#,##0.00", "0.00%"], ["m/d/yyyy", "General"]])</returns>
    Task<RangeNumberFormatResult> GetNumberFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Sets uniform number format for entire range
    /// Excel COM: Range.NumberFormat = formatCode
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <param name="formatCode">
    /// Excel format code (e.g., "$#,##0.00", "0.00%", "m/d/yyyy", "General", "@")
    /// See NumberFormatPresets class for common patterns
    /// </param>
    Task<OperationResult> SetNumberFormatAsync(IExcelBatch batch, string sheetName, string rangeAddress, string formatCode);

    /// <summary>
    /// Sets number formats cell-by-cell from 2D array
    /// Excel COM: Range.NumberFormat (per cell)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <param name="formats">2D array of format codes matching range dimensions</param>
    Task<OperationResult> SetNumberFormatsAsync(IExcelBatch batch, string sheetName, string rangeAddress, List<List<string>> formats);

    // === FORMATTING OPERATIONS ===

    /// <summary>
    /// Applies built-in Excel cell style to range (recommended for consistency)
    /// Excel COM: Range.Style = styleName
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Sheet name (empty for active sheet)</param>
    /// <param name="rangeAddress">Range address (e.g., "A1:D10")</param>
    /// <param name="styleName">
    /// Built-in style name (e.g., "Heading 1", "Accent1", "Good", "Total", "Currency", "Percent")
    /// Use "Normal" to reset to default formatting
    /// </param>
    /// <remarks>
    /// Built-in styles are theme-aware and provide professional, consistent formatting.
    /// Common styles: Heading 1-4, Title, Total, Input, Output, Calculation,
    /// Good/Bad/Neutral, Accent1-6, Note, Warning, Currency, Percent, Comma
    /// </remarks>
    Task<OperationResult> SetStyleAsync(IExcelBatch batch, string sheetName, string rangeAddress, string styleName);

    /// <summary>
    /// Gets the current built-in style name applied to a range
    /// Excel COM: Range.Style.Name property
    /// </summary>
    /// <param name="batch">Excel batch context</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range address (e.g., "A1" for single cell, "A1:D10" for range)</param>
    /// <returns>RangeStyleResult with current style name and whether it's a built-in style</returns>
    /// <remarks>
    /// Returns the style name of the first cell in the range.
    /// Use this to inspect current formatting before applying changes.
    /// </remarks>
    Task<RangeStyleResult> GetStyleAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Applies visual formatting to range (font, fill, border, alignment)
    /// Excel COM: Range.Font, Range.Interior, Range.Borders, Range.HorizontalAlignment, etc.
    /// </summary>
    /// <remarks>
    /// For consistent, professional formatting, prefer SetStyleAsync() with built-in styles.
    /// Use FormatRangeAsync() only when built-in styles don't meet your needs.
    /// </remarks>
    Task<OperationResult> FormatRangeAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
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
    /// Adds data validation rules to range
    /// Excel COM: Range.Validation.Add()
    /// </summary>
    Task<OperationResult> ValidateRangeAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string validationType,
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
    /// Gets data validation settings from first cell in range
    /// Excel COM: Range.Validation
    /// </summary>
    Task<RangeValidationResult> GetValidationAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Removes data validation from range
    /// Excel COM: Range.Validation.Delete()
    /// </summary>
    Task<OperationResult> RemoveValidationAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    // === AUTO-FIT OPERATIONS ===

    /// <summary>
    /// Auto-fits column widths to content
    /// Excel COM: Range.Columns.AutoFit()
    /// </summary>
    Task<OperationResult> AutoFitColumnsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Auto-fits row heights to content
    /// Excel COM: Range.Rows.AutoFit()
    /// </summary>
    Task<OperationResult> AutoFitRowsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    // === MERGE OPERATIONS ===

    /// <summary>
    /// Merges cells in range into a single cell
    /// Excel COM: Range.Merge()
    /// </summary>
    Task<OperationResult> MergeCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Unmerges previously merged cells
    /// Excel COM: Range.UnMerge()
    /// </summary>
    Task<OperationResult> UnmergeCellsAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    /// <summary>
    /// Checks if range contains merged cells
    /// Excel COM: Range.MergeCells
    /// </summary>
    Task<RangeMergeInfoResult> GetMergeInfoAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    // === CONDITIONAL FORMATTING OPERATIONS ===

    /// <summary>
    /// Adds conditional formatting rule to range
    /// Excel COM: Range.FormatConditions.Add()
    /// </summary>
    Task<OperationResult> AddConditionalFormattingAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string ruleType,
        string? formula1,
        string? formula2,
        string? formatStyle);

    /// <summary>
    /// Removes all conditional formatting from range
    /// Excel COM: Range.FormatConditions.Delete()
    /// </summary>
    Task<OperationResult> ClearConditionalFormattingAsync(IExcelBatch batch, string sheetName, string rangeAddress);

    // === CELL PROTECTION OPERATIONS ===

    /// <summary>
    /// Locks or unlocks cells (requires worksheet protection to take effect)
    /// Excel COM: Range.Locked
    /// </summary>
    Task<OperationResult> SetCellLockAsync(IExcelBatch batch, string sheetName, string rangeAddress, bool locked);

    /// <summary>
    /// Gets lock status of first cell in range
    /// Excel COM: Range.Locked
    /// </summary>
    Task<RangeLockInfoResult> GetCellLockAsync(IExcelBatch batch, string sheetName, string rangeAddress);
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
