using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable analysis operations (GetData, SetFieldFilter, SortField)
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Gets the current data from a PivotTable
    /// </summary>
    public PivotTableDataResult GetData(IExcelBatch batch, string pivotTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? tableRange = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                tableRange = pivot.TableRange2;

                // Get the data range
                object[,] values = tableRange.Value2;

                // Convert to List<List<object?>>
                var dataList = new List<List<object?>>();
                for (int row = 1; row <= values.GetLength(0); row++)
                {
                    var rowList = new List<object?>();
                    for (int col = 1; col <= values.GetLength(1); col++)
                    {
                        rowList.Add(values[row, col]);
                    }
                    dataList.Add(rowList);
                }

                return new PivotTableDataResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    Values = dataList,
                    DataRowCount = values.GetLength(0),
                    DataColumnCount = values.GetLength(1),
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotTableDataResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to get PivotTable data: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Sets filter for a field
    /// </summary>
    public PivotFieldFilterResult SetFieldFilter(IExcelBatch batch, string pivotTableName,
        string fieldName, List<string> selectedValues)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.SetFieldFilter(pivot, fieldName, selectedValues, batch.WorkbookPath);
            }
            catch (Exception ex)
            {
                return new PivotFieldFilterResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to set field filter: {ex.Message}",
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
    /// Sorts a field
    /// </summary>
    public PivotFieldResult SortField(IExcelBatch batch, string pivotTableName,
        string fieldName, SortDirection direction = SortDirection.Ascending)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);

                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.SortField(pivot, fieldName, direction, batch.WorkbookPath);
            }
            catch (Exception ex)
            {
                return new PivotFieldResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to sort field: {ex.Message}",
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

