using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

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

                var values = ExcelValueNormalizer.Normalize(tableRange.Value2);

                return new PivotTableDataResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    Values = values.Values,
                    DataRowCount = values.RowCount,
                    DataColumnCount = values.ColumnCount,
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
        => ExecuteWithStrategy<PivotFieldFilterResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.SetFieldFilter(pivot, fieldName, selectedValues, batch.WorkbookPath));

    /// <summary>
    /// Sets an OLAP report filter on exactly one worksheet PivotTable.
    /// An empty selection means all members.
    /// </summary>
    public PivotReportFilterResult SetReportFilter(IExcelBatch batch, string sheetName,
        string pivotTableName, string fieldName, List<string> selectedItems)
    {
        var invalidRequest = ValidateReportFilterRequest(
            batch.WorkbookPath, sheetName, pivotTableName, fieldName, selectedItems);
        if (invalidRequest is not null)
        {
            return invalidRequest;
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivotTables = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubeFields = null;
            Excel.CubeField? cubeField = null;
            Excel.PivotFields? pivotFields = null;
            Excel.PivotField? pivotField = null;
            var result = new PivotReportFilterResult
            {
                SheetName = sheetName,
                PivotTableName = pivotTableName,
                FieldName = fieldName,
                FilePath = batch.WorkbookPath
            };
            var filterMutationStarted = false;
            var manualUpdateChanged = false;
            var previousManualUpdate = false;

            try
            {
                ct.ThrowIfCancellationRequested();
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                pivotTables = (Excel.PivotTables)sheet.PivotTables();
                pivot = pivotTables.Item(pivotTableName);
                if (!PivotTableHelpers.IsOlapPivotTable(pivot))
                {
                    throw new InvalidOperationException(
                        $"PivotTable '{pivotTableName}' on worksheet '{sheetName}' is not an OLAP/Data Model PivotTable.");
                }

                cubeFields = pivot.CubeFields;
                cubeField = cubeFields[fieldName];
                if (Convert.ToInt32(cubeField.Orientation,
                        System.Globalization.CultureInfo.InvariantCulture) != XlPivotFieldOrientation.xlPageField)
                {
                    throw new InvalidOperationException(
                        $"OLAP field '{fieldName}' must be in the report-filter area before its selection can be changed.");
                }

                pivotFields = cubeField.PivotFields;
                if (pivotFields.Count == 0)
                {
                    throw new InvalidOperationException(
                        $"OLAP report-filter field '{fieldName}' has no PivotField available.");
                }
                pivotField = pivotFields.Item(1);

                previousManualUpdate = pivot.ManualUpdate;
                if (!previousManualUpdate)
                {
                    pivot.ManualUpdate = true;
                    manualUpdateChanged = true;
                }

                filterMutationStarted = true;
                if (selectedItems.Count == 0)
                {
                    pivotField.ClearAllFilters();
                }
                else if (selectedItems.Count == 1)
                {
                    cubeField.EnableMultiplePageItems = false;
                    pivotField.CurrentPageName = selectedItems[0];
                }
                else
                {
                    cubeField.EnableMultiplePageItems = true;
                    pivotField.VisibleItemsList = selectedItems.ToArray();
                }

                ReadReportFilterSelection(cubeField, pivotField, result);
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage =
                    $"Failed to set report filter '{fieldName}' on PivotTable '{pivotTableName}' " +
                    $"in worksheet '{sheetName}': {ex.Message}";
                result.MayHavePartiallyChanged = filterMutationStarted;
                if (filterMutationStarted && cubeField is not null && pivotField is not null)
                {
                    try
                    {
                        ReadReportFilterSelection(cubeField, pivotField, result);
                    }
                    catch (Exception readException)
                    {
                        result.ErrorMessage +=
                            $" Actual filter selection could not be read: {readException.Message}";
                    }
                }
            }
            finally
            {
                if (manualUpdateChanged && pivot is not null)
                {
                    try
                    {
                        pivot.ManualUpdate = previousManualUpdate;
                    }
                    catch (Exception restoreException)
                    {
                        result.Success = false;
                        result.MayHavePartiallyChanged = true;
                        result.ErrorMessage = string.IsNullOrEmpty(result.ErrorMessage)
                            ? $"Could not restore PivotTable update behavior: {restoreException.Message}"
                            : $"{result.ErrorMessage} Could not restore PivotTable update behavior: {restoreException.Message}";
                    }
                }

                ComUtilities.Release(ref pivotField);
                ComUtilities.Release(ref pivotFields);
                ComUtilities.Release(ref cubeField);
                ComUtilities.Release(ref cubeFields);
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref pivotTables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }

            return result;
        });
    }

    private static PivotReportFilterResult? ValidateReportFilterRequest(
        string workbookPath, string sheetName, string pivotTableName, string fieldName,
        List<string> selectedItems)
    {
        string? error = null;
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            error = "Worksheet name must not be empty.";
        }
        else if (string.IsNullOrWhiteSpace(pivotTableName))
        {
            error = "PivotTable name must not be empty.";
        }
        else if (string.IsNullOrWhiteSpace(fieldName))
        {
            error = "Field name must not be empty.";
        }
        else if (selectedItems is null)
        {
            error = "Selected items must be provided; use an empty list to show all members.";
        }
        else if (selectedItems.Any(string.IsNullOrWhiteSpace))
        {
            error = "Selected member names must not be empty.";
        }
        else if (selectedItems.Count != selectedItems.Distinct(StringComparer.Ordinal).Count())
        {
            error = "Selected member names must not contain duplicates.";
        }

        return error is null
            ? null
            : new PivotReportFilterResult
            {
                Success = false,
                ErrorMessage = error,
                SheetName = sheetName,
                PivotTableName = pivotTableName,
                FieldName = fieldName,
                FilePath = workbookPath
            };
    }

    private static void ReadReportFilterSelection(
        Excel.CubeField cubeField, Excel.PivotField pivotField,
        PivotReportFilterResult result)
    {
        if (cubeField.EnableMultiplePageItems)
        {
            if (cubeField.AllItemsVisible)
            {
                result.ShowAll = true;
                result.SelectedItems.Clear();
                return;
            }

            var visibleItems = pivotField.VisibleItemsList;
            result.SelectedItems = visibleItems is Array items
                ? items.Cast<object>().Select(item => item.ToString() ?? string.Empty).ToList()
                : [];
        }
        else
        {
            var currentPageName = pivotField.CurrentPageName;
            result.ShowAll = cubeField.AllItemsVisible &&
                (string.IsNullOrEmpty(currentPageName) || currentPageName == "(All)");
            result.SelectedItems = result.ShowAll ? [] : [currentPageName];
        }

    }

    /// <summary>
    /// Sorts a field
    /// </summary>
    public PivotFieldResult SortField(IExcelBatch batch, string pivotTableName,
        string fieldName, SortDirection direction = SortDirection.Ascending)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.SortField(pivot, fieldName, direction, batch.WorkbookPath));
}
