using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range edit operations - insert, delete, find, replace, sort rows/columns.
/// Use range command for values/formulas/copy/clear operations.
/// </summary>
[ServiceCategory("rangeedit", "RangeEdit")]
[McpTool("excel_range_edit")]
public interface IRangeEditCommands
{
    // === INSERT/DELETE CELL OPERATIONS ===

    /// <summary>
    /// Inserts blank cells, shifting existing cells down or right.
    /// Excel COM: Range.Insert(shift)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range where cells will be inserted</param>
    /// <param name="insertShift">Direction to shift existing cells (Down or Right)</param>
    [ServiceAction("insert-cells")]
    OperationResult InsertCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] InsertShiftDirection insertShift);

    /// <summary>
    /// Deletes cells, shifting remaining cells up or left.
    /// Excel COM: Range.Delete(shift)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range to delete</param>
    /// <param name="deleteShift">Direction to shift remaining cells (Up or Left)</param>
    [ServiceAction("delete-cells")]
    OperationResult DeleteCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] DeleteShiftDirection deleteShift);

    /// <summary>
    /// Inserts entire rows above the range.
    /// Excel COM: Range.EntireRow.Insert()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range defining rows to insert above</param>
    [ServiceAction("insert-rows")]
    OperationResult InsertRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Deletes entire rows in the range.
    /// Excel COM: Range.EntireRow.Delete()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range defining rows to delete</param>
    [ServiceAction("delete-rows")]
    OperationResult DeleteRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Inserts entire columns to the left of the range.
    /// Excel COM: Range.EntireColumn.Insert()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range defining columns to insert left of</param>
    [ServiceAction("insert-columns")]
    OperationResult InsertColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Deletes entire columns in the range.
    /// Excel COM: Range.EntireColumn.Delete()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range defining columns to delete</param>
    [ServiceAction("delete-columns")]
    OperationResult DeleteColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === FIND/REPLACE OPERATIONS ===

    /// <summary>
    /// Finds all cells matching criteria in range.
    /// Excel COM: Range.Find()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range to search in</param>
    /// <param name="searchValue">Value to find</param>
    /// <param name="findOptions">Search options (case, whole cell, etc.)</param>
    [ServiceAction("find")]
    RangeFindResult Find(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string searchValue, FindOptions findOptions);

    /// <summary>
    /// Replaces text/values in range.
    /// Excel COM: Range.Replace()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range to search in</param>
    /// <param name="findValue">Value to find</param>
    /// <param name="replaceValue">Value to replace with</param>
    /// <param name="replaceOptions">Replace options (case, whole cell, etc.)</param>
    [ServiceAction("replace")]
    void Replace(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string findValue, [RequiredParameter] string replaceValue, ReplaceOptions replaceOptions);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts range by one or more columns.
    /// Excel COM: Range.Sort()
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="sheetName">Worksheet name</param>
    /// <param name="rangeAddress">Range to sort</param>
    /// <param name="sortColumns">Columns to sort by with direction</param>
    /// <param name="hasHeaders">True if first row contains headers</param>
    [ServiceAction("sort")]
    void Sort(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] List<SortColumn> sortColumns, bool hasHeaders = true);
}
