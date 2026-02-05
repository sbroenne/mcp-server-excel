using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable grand totals commands - show/hide row and column grand totals.
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Shows or hides grand totals for rows and/or columns in the PivotTable.
    /// </summary>
    /// <param name="batch">The Excel batch session containing the workbook.</param>
    /// <param name="pivotTableName">Name of the PivotTable to configure.</param>
    /// <param name="showRowGrandTotals">Show row grand totals (bottom summary row).</param>
    /// <param name="showColumnGrandTotals">Show column grand totals (right summary column).</param>
    /// <returns>Operation result indicating success or failure.</returns>
    /// <remarks>
    /// GRAND TOTALS:
    /// - Row Grand Totals: Summary row displayed at the bottom of the PivotTable
    /// - Column Grand Totals: Summary column displayed at the right of the PivotTable
    /// - Independent control: Can show/hide row and column totals separately
    ///
    /// SUPPORT:
    /// - Regular PivotTables: Full support
    /// - OLAP PivotTables: Full support (same COM properties)
    /// </remarks>
    public OperationResult SetGrandTotals(IExcelBatch batch, string pivotTableName, bool showRowGrandTotals, bool showColumnGrandTotals)
        => ExecuteWithStrategy<OperationResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.SetGrandTotals(pivot, showRowGrandTotals, showColumnGrandTotals, batch.WorkbookPath, batch.Logger));
}


