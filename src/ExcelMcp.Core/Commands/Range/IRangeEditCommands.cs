using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Excel range edit operations - insert, delete, find, replace, sort.
/// Use IRangeCommands for values/formulas/copy/clear operations.
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
    [ServiceAction("insert-cells")]
    OperationResult InsertCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] InsertShiftDirection insertShift);

    /// <summary>
    /// Deletes cells, shifting remaining cells up or left.
    /// Excel COM: Range.Delete(shift)
    /// </summary>
    [ServiceAction("delete-cells")]
    OperationResult DeleteCells(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] DeleteShiftDirection deleteShift);

    /// <summary>
    /// Inserts entire rows above the range.
    /// Excel COM: Range.EntireRow.Insert()
    /// </summary>
    [ServiceAction("insert-rows")]
    OperationResult InsertRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Deletes entire rows in the range.
    /// Excel COM: Range.EntireRow.Delete()
    /// </summary>
    [ServiceAction("delete-rows")]
    OperationResult DeleteRows(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Inserts entire columns to the left of the range.
    /// Excel COM: Range.EntireColumn.Insert()
    /// </summary>
    [ServiceAction("insert-columns")]
    OperationResult InsertColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    /// <summary>
    /// Deletes entire columns in the range.
    /// Excel COM: Range.EntireColumn.Delete()
    /// </summary>
    [ServiceAction("delete-columns")]
    OperationResult DeleteColumns(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress);

    // === FIND/REPLACE OPERATIONS ===

    /// <summary>
    /// Finds all cells matching criteria in range.
    /// Excel COM: Range.Find()
    /// </summary>
    [ServiceAction("find")]
    RangeFindResult Find(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string searchValue, FindOptions findOptions);

    /// <summary>
    /// Replaces text/values in range.
    /// Excel COM: Range.Replace()
    /// </summary>
    [ServiceAction("replace")]
    void Replace(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] string findValue, [RequiredParameter] string replaceValue, ReplaceOptions replaceOptions);

    // === SORT OPERATIONS ===

    /// <summary>
    /// Sorts range by one or more columns.
    /// Excel COM: Range.Sort()
    /// </summary>
    [ServiceAction("sort")]
    void Sort(IExcelBatch batch, string sheetName, [RequiredParameter] string rangeAddress, [RequiredParameter] List<SortColumn> sortColumns, bool hasHeaders = true);
}
