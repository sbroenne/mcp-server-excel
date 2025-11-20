using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable grouping operations (GroupByDate, GroupByNumeric)
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Groups a date/time field by the specified interval (Month, Quarter, Year)
    /// </summary>
    public PivotFieldResult GroupByDate(IExcelBatch batch, string pivotTableName,
        string fieldName, DateGroupingInterval interval)
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
                return strategy.GroupByDate(pivot, fieldName, interval, batch.WorkbookPath, batch.Logger);
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to group field '{fieldName}' by date: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Groups a numeric field by the specified interval (e.g., 0-100, 100-200, 200-300)
    /// </summary>
    public PivotFieldResult GroupByNumeric(IExcelBatch batch, string pivotTableName,
        string fieldName, double? start, double? endValue, double intervalSize)
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
                return strategy.GroupByNumeric(pivot, fieldName, start, endValue, intervalSize, batch.WorkbookPath, batch.Logger);
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to group field '{fieldName}' numerically: {ex.Message}",
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
