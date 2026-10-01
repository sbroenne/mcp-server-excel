using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range editing operations: insert/delete cells, rows, and columns; find/replace text; sort data.
/// Use range for values/formulas/copy/clear operations.
///
/// INSERT/DELETE CELLS: Specify shift direction to control how surrounding cells move.
/// - Insert: 'Down' or 'Right'
/// - Delete: 'Up' or 'Left'
///
/// INSERT/DELETE ROWS: Use row range like '5:10' to insert/delete rows 5-10.
/// INSERT/DELETE COLUMNS: Use column range like 'B:D' to insert/delete columns B-D.
///
/// FIND/REPLACE: Search within the specified range with optional case/cell matching.
/// - Find returns up to maxMatches cells (default: 10), with exact totalCount, returnedCount, and truncated.
/// - Exact counting searches all matches, even after the returned-cell limit is reached.
/// - Replace modifies all matches by default.
///
/// SORT: Specify sortColumns as array of {columnIndex: 1, ascending: true} objects.
/// Column indices are 1-based relative to the range.
/// </summary>
[ServiceCategory("rangeedit", "RangeEdit")]
[McpTool("range_edit", Title = "Range Edit Operations", Destructive = true, Category = "data",
    Description = "Insert/delete cells, rows, or columns; find/replace text; sort data. Cell movement uses insertShift (Down/Right) or deleteShift (Up/Left). Rows use a range like 5:10; columns use B:D. Find returns up to max_matches cells (default: 10) with exact totalCount, returnedCount, truncated, and optional case/cell matching. Exact counting searches all matches; the limit bounds returned cell details, not search time. Replace modifies all matches by default (replace_options.replaceAll=true). Sort uses sortColumns, an array of {columnIndex, ascending}; indices are 1-based relative to the range.")]
public interface IRangeEditCommands
{
    // === INSERT/DELETE CELL OPERATIONS ===

    /// <summary>
    /// Inserts blank cells, shifting existing cells down or right.
    /// Excel COM: Range.Insert(shift)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address where cells will be inserted (e.g., 'A1:D10')</param>
    /// <param name="insertShift">Direction to shift existing cells: 'Down' or 'Right'</param>
    [ServiceAction("insert-cells")]
    OperationResult InsertCells(
        IExcelBatch batch, string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter]
        [FromString] InsertShiftDirection insertShift);

    /// <summary>
    /// Deletes cells, shifting remaining cells up or left.
    /// Excel COM: Range.Delete(shift)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to delete (e.g., 'A1:D10')</param>
    /// <param name="deleteShift">Direction to shift remaining cells: 'Up' or 'Left'</param>
    [ServiceAction("delete-cells")]
    OperationResult DeleteCells(
        IExcelBatch batch, string sheetName,
        [RequiredParameter] string rangeAddress,
        [RequiredParameter]
        [FromString] DeleteShiftDirection deleteShift);

    /// <summary>
    /// Inserts entire rows above the range.
    /// Excel COM: Range.EntireRow.Insert()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row range defining rows to insert above (e.g., '5:10' for rows 5-10)</param>
    [ServiceAction("insert-rows")]
    OperationResult InsertRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Deletes entire rows in the range.
    /// Excel COM: Range.EntireRow.Delete()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Row range defining rows to delete (e.g., '5:10' for rows 5-10)</param>
    [ServiceAction("delete-rows")]
    OperationResult DeleteRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Inserts entire columns to the left of the range.
    /// Excel COM: Range.EntireColumn.Insert()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Column range defining columns to insert left of (e.g., 'B:D' for columns B-D)</param>
    [ServiceAction("insert-columns")]
    OperationResult InsertColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Deletes entire columns in the range.
    /// Excel COM: Range.EntireColumn.Delete()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="rangeAddress">Column range defining columns to delete (e.g., 'B:D' for columns B-D)</param>
    [ServiceAction("delete-columns")]
    OperationResult DeleteColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === FIND/REPLACE OPERATIONS ===

    /// <summary>
    /// Counts all matching cells and returns up to maxMatches cell details.
    /// Excel COM: Range.Find()/FindNext(). Exact counting traverses every match.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to search within (e.g., 'A1:D100')</param>
    /// <param name="searchValue">Text or value to search for</param>
    /// <param name="findOptions">Search options: matchCase (default: false), matchEntireCell (default: false), searchFormulas (default: true)</param>
    /// <param name="maxMatches">Maximum matching cells to return (default: 10). Positive whole number from 1 through 2147483647. All matches are still counted.</param>
    [ServiceAction("find")]
    RangeFindResult Find(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string searchValue, FindOptions findOptions, int maxMatches = 10);

    /// <summary>
    /// Replaces text/values in range.
    /// Excel COM: Range.Replace()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to search within (e.g., 'A1:D100')</param>
    /// <param name="findValue">Text or value to search for</param>
    /// <param name="replaceValue">Text or value to replace matches with</param>
    /// <param name="replaceOptions">Replace options: matchCase (default: false), matchEntireCell (default: false), replaceAll (default: true)</param>
    [ServiceAction("replace")]
    OperationResult Replace(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string findValue, [RequiredParameter] string replaceValue, ReplaceOptions replaceOptions);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts range by one or more columns.
    /// Excel COM: Range.Sort()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Name of the worksheet containing the range</param>
    /// <param name="rangeAddress">Cell range address to sort (e.g., 'A1:D100')</param>
    /// <param name="sortColumns">Array of sort specifications: [{columnIndex: 1, ascending: true}, ...] - columnIndex is 1-based relative to range</param>
    /// <param name="hasHeaders">Whether the range has a header row to exclude from sorting (default: true)</param>
    [ServiceAction("sort")]
    OperationResult Sort(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] List<SortColumn> sortColumns, bool hasHeaders = true);
}
