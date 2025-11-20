using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop;
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
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                if (pivot == null)
                {
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"PivotTable '{pivotTableName}' not found",
                        FilePath = batch.WorkbookPath
                    };
                }

                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.SetGrandTotals(pivot, showRowGrandTotals, showColumnGrandTotals, batch.WorkbookPath, batch.Logger);
            }
            catch (Exception ex)
            {
#pragma warning disable CA1848
                batch.Logger?.LogError(ex, "SetGrandTotals failed for PivotTable {PivotTableName}", pivotTableName);
#pragma warning restore CA1848
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to set grand totals: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }
}
