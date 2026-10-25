using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable management commands - main partial class with shared state and helper methods
/// </summary>
public partial class PivotTableCommands : IPivotTableCommands, IPivotTableFieldCommands, IPivotTableCalcCommands
{
    #region Helper Methods

    /// <summary>
    /// Finds a PivotTable by name in the workbook.
    /// Delegates to CoreLookupHelpers.FindPivotTable for the actual lookup.
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="pivotTableName">Name of the PivotTable to find</param>
    /// <returns>The PivotTable object if found</returns>
    /// <exception cref="InvalidOperationException">Thrown if PivotTable is not found</exception>
    private static Excel.PivotTable FindPivotTable(Excel.Workbook workbook, string pivotTableName)
        => CoreLookupHelpers.FindPivotTable(workbook, pivotTableName);

    /// <summary>
    /// Executes a strategy-based operation on a PivotTable.
    /// Centralizes the common pattern: find pivot → get strategy → execute → release.
    /// </summary>
    /// <typeparam name="TResult">The result type returned by the strategy operation</typeparam>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="operation">The strategy operation to execute (receives strategy and pivot)</param>
    /// <returns>The result from the strategy operation</returns>
    private static TResult ExecuteWithStrategy<TResult>(
        IExcelBatch batch,
        string pivotTableName,
        Func<IPivotTableFieldStrategy, Excel.PivotTable, TResult> operation)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return operation(strategy, pivot);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    #endregion
}
