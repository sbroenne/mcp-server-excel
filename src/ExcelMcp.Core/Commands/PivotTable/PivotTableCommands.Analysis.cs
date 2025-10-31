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
    public async Task<PivotTableDataResult> GetDataAsync(IExcelBatch batch, string pivotTableName)
    {
        return await batch.Execute((ctx, ct) =>
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
    public async Task<PivotFieldFilterResult> SetFieldFilterAsync(IExcelBatch batch, string pivotTableName,
        string fieldName, List<string> selectedValues)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;
            dynamic? pivotItems = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                field = pivot.PivotFields.Item(fieldName);
                pivotItems = field.PivotItems;

                // First, show selected items (must do this before hiding others to avoid Excel error)
                var selectedSet = new HashSet<string>(selectedValues);
                foreach (string value in selectedValues)
                {
                    dynamic? item = null;
                    try
                    {
                        item = pivotItems.Item(value);
                        item.Visible = true;
                    }
                    catch
                    {
                        // Item not found, skip
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }

                // Then, hide unselected items (never hides the last item because we showed selected ones first)
                for (int i = 1; i <= pivotItems.Count; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = pivotItems.Item(i);
                        string itemName = item.Name?.ToString() ?? string.Empty;
                        if (!selectedSet.Contains(itemName))
                        {
                            item.Visible = false;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }

                // Refresh
                pivot.RefreshTable();

                // Get all available items
                var availableItems = new List<string>();
                for (int i = 1; i <= pivotItems.Count; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = pivotItems.Item(i);
                        availableItems.Add(item.Name?.ToString() ?? string.Empty);
                    }
                    finally
                    {
                        ComUtilities.Release(ref item);
                    }
                }

                return new PivotFieldFilterResult
                {
                    Success = true,
                    FieldName = fieldName,
                    SelectedItems = selectedValues,
                    AvailableItems = availableItems,
                    ShowAll = false,
                    FilePath = batch.WorkbookPath
                };
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
                ComUtilities.Release(ref pivotItems);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Sorts a field
    /// </summary>
    public async Task<PivotFieldResult> SortFieldAsync(IExcelBatch batch, string pivotTableName,
        string fieldName, SortDirection direction = SortDirection.Ascending)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? field = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                field = pivot.PivotFields.Item(fieldName);

                // Excel sort order: xlAscending = 1, xlDescending = 2
                int sortOrder = direction == SortDirection.Ascending ? 1 : 2;

                // AutoSort method: xlManual = -4135, xlAscending = 1, xlDescending = 2
                field.AutoSort(sortOrder, fieldName);

                // Refresh
                pivot.RefreshTable();

                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = fieldName,
                    CustomName = field.Caption?.ToString() ?? fieldName,
                    Area = (PivotFieldArea)field.Orientation,
                    FilePath = batch.WorkbookPath
                };
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
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref pivot);
            }
        });
    }
}
