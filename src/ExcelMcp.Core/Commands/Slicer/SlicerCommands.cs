using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

/// <summary>
/// Slicer commands bridging PivotTable and Table slicer operations.
/// </summary>
public sealed partial class SlicerCommands : ISlicerCommands
{
    /// <inheritdoc />
    public SlicerResult CreateSlicer(IExcelBatch batch, string pivotTableName,
        string fieldName, string slicerName, string destinationSheet, string position)
        => PivotTableCommands.CreateSlicer(batch, pivotTableName, fieldName, slicerName, destinationSheet, position);

    /// <inheritdoc />
    public SlicerListResult ListSlicers(IExcelBatch batch, string? pivotTableName = null)
        => PivotTableCommands.ListSlicers(batch, pivotTableName);

    /// <inheritdoc />
    public SlicerResult SetSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true)
        => PivotTableCommands.SetSlicerSelection(batch, slicerName, selectedItems, clearFirst);

    /// <inheritdoc />
    public OperationResult DeleteSlicer(IExcelBatch batch, string slicerName)
        => PivotTableCommands.DeleteSlicer(batch, slicerName);

    /// <inheritdoc />
    public SlicerResult CreateTableSlicer(IExcelBatch batch, string tableName,
        string columnName, string slicerName, string destinationSheet, string position)
        => TableCommands.CreateTableSlicer(batch, tableName, columnName, slicerName, destinationSheet, position);

    /// <inheritdoc />
    public SlicerListResult ListTableSlicers(IExcelBatch batch, string? tableName = null)
        => TableCommands.ListTableSlicers(batch, tableName);

    /// <inheritdoc />
    public SlicerResult SetTableSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true)
        => TableCommands.SetTableSlicerSelection(batch, slicerName, selectedItems, clearFirst);

    /// <inheritdoc />
    public OperationResult DeleteTableSlicer(IExcelBatch batch, string slicerName)
        => TableCommands.DeleteTableSlicer(batch, slicerName);
}
