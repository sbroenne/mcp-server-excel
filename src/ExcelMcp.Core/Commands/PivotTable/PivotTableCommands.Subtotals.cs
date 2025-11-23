using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable subtotals operations - show/hide automatic subtotals for row fields.
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Shows or hides subtotals for a specific row field.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="pivotTableName">Name of the PivotTable</param>
    /// <param name="fieldName">Name of the row field</param>
    /// <param name="showSubtotals">True to show automatic subtotals, false to hide</param>
    /// <returns>Result with updated field configuration</returns>
    /// <remarks>
    /// SUBTOTALS BEHAVIOR:
    /// - When enabled: Shows Automatic subtotals (uses appropriate function based on data)
    /// - When disabled: Hides all subtotals for the field
    ///
    /// OLAP LIMITATION:
    /// - OLAP PivotTables only support Automatic subtotals
    /// - Regular PivotTables can choose Sum, Count, Average, etc. (future enhancement)
    /// </remarks>
    public PivotFieldResult SetSubtotals(IExcelBatch batch, string pivotTableName, string fieldName, bool showSubtotals)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.SetSubtotals(pivot, fieldName, showSubtotals, batch.WorkbookPath, batch.Logger);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }
}
