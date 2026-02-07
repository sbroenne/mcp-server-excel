using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

/// <summary>
/// Slicer commands bridging PivotTable and Table slicer operations.
/// </summary>
public sealed class SlicerCommands : ISlicerCommands
{
    private readonly PivotTableCommands _pivotTableCommands = new();
    private readonly TableCommands _tableCommands = new();

    /// <inheritdoc />
    public SlicerResult CreateSlicer(IExcelBatch batch, string pivotTableName,
        string fieldName, string slicerName, string destinationSheet, string position)
        => _pivotTableCommands.CreateSlicer(batch, pivotTableName, fieldName, slicerName, destinationSheet, position);

    /// <inheritdoc />
    public SlicerListResult ListSlicers(IExcelBatch batch, string? pivotTableName = null)
        => _pivotTableCommands.ListSlicers(batch, pivotTableName);

    /// <inheritdoc />
    public SlicerResult SetSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true)
        => _pivotTableCommands.SetSlicerSelection(batch, slicerName, selectedItems, clearFirst);

    /// <inheritdoc />
    public OperationResult DeleteSlicer(IExcelBatch batch, string slicerName)
        => _pivotTableCommands.DeleteSlicer(batch, slicerName);

    /// <inheritdoc />
    public SlicerResult CreateTableSlicer(IExcelBatch batch, string tableName,
        string columnName, string slicerName, string destinationSheet, string position)
        => _tableCommands.CreateTableSlicer(batch, tableName, columnName, slicerName, destinationSheet, position);

    /// <inheritdoc />
    public SlicerListResult ListTableSlicers(IExcelBatch batch, string? tableName = null)
        => _tableCommands.ListTableSlicers(batch, tableName);

    /// <inheritdoc />
    public SlicerResult SetTableSlicerSelection(IExcelBatch batch, string slicerName,
        List<string> selectedItems, bool clearFirst = true)
        => _tableCommands.SetTableSlicerSelection(batch, slicerName, selectedItems, clearFirst);

    /// <inheritdoc />
    public OperationResult DeleteTableSlicer(IExcelBatch batch, string slicerName)
        => _tableCommands.DeleteTableSlicer(batch, slicerName);
}
