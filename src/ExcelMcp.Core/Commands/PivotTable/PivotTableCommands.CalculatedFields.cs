using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Calculated Fields operations for PivotTableCommands.
/// Creates custom fields with formulas for Regular PivotTables.
/// OLAP PivotTables use DAX measures instead (see excel_datamodel tool).
/// </summary>
public partial class PivotTableCommands
{
    /// <inheritdoc/>
    public PivotFieldResult CreateCalculatedField(IExcelBatch batch, string pivotTableName,
        string fieldName, string formula)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Determine strategy (OLAP vs Regular)
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);

                // Delegate to strategy with logger
                return strategy.CreateCalculatedField(pivot, fieldName, formula, batch.WorkbookPath, batch.Logger);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }
}
