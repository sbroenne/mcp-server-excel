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
            Excel.PivotTable? pivot = null;
            Excel.PivotFields? rowFields = null;
            Excel.Range? tableRange = null;
            Excel.Range? reportRange = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                rowFields = (Excel.PivotFields)pivot.RowFields;
                tableRange = pivot.TableRange2;
                reportRange = pivot.TableRange1;

                var values = ExcelValueNormalizer.Normalize((object?)tableRange.Value2);
                var report = ExcelValueNormalizer.Normalize((object?)reportRange.Value2);
                var headerRow = report.Values.FirstOrDefault() ?? [];
                bool hasRowFields = rowFields.Count > 0;
                bool hasGrandTotalRow = hasRowFields && pivot.ColumnGrand && report.RowCount > 1;
                int firstDataColumn = hasRowFields ? 1 : 0;
                var columnHeaders = headerRow
                    .Skip(firstDataColumn)
                    .Select(value => value?.ToString() ?? string.Empty)
                    .ToList();
                int lastDataRow = report.RowCount;
                if (hasGrandTotalRow)
                    lastDataRow--;

                var rowHeaders = hasRowFields
                    ? report.Values
                        .Skip(1)
                        .Take(lastDataRow - 1)
                        .Select(row => row.Count > 0 ? row[0]?.ToString() ?? string.Empty : string.Empty)
                        .ToList()
                    : [];
                var grandTotals = new Dictionary<string, object?>(StringComparer.Ordinal);
                if (hasGrandTotalRow)
                {
                    var totalRow = report.Values[^1];
                    for (int columnIndex = firstDataColumn;
                         columnIndex < headerRow.Count && columnIndex < totalRow.Count;
                         columnIndex++)
                    {
                        var header = headerRow[columnIndex]?.ToString();
                        if (!string.IsNullOrWhiteSpace(header))
                            grandTotals[header] = totalRow[columnIndex];
                    }
                }

                return new PivotTableDataResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    Values = values.Values,
                    ColumnHeaders = columnHeaders,
                    RowHeaders = rowHeaders,
                    DataRowCount = values.RowCount,
                    DataColumnCount = values.ColumnCount,
                    GrandTotals = grandTotals,
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref reportRange);
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref rowFields);
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

    /// <summary>Reads an OLAP report filter from exactly one worksheet PivotTable.</summary>
    public PivotReportFilterResult GetReportFilter(
        IExcelBatch batch, string sheetName, string pivotTableName, string fieldName)
    {
        var validationError = ValidateReportFilterTarget(sheetName, pivotTableName, fieldName);
        var result = new PivotReportFilterResult
        {
            SheetName = sheetName,
            PivotTableName = pivotTableName,
            FieldName = fieldName,
            FilePath = batch.WorkbookPath
        };
        if (validationError is not null)
        {
            result.ErrorMessage = validationError;
            return result;
        }

        return batch.Execute(ReadReportFilter);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "Reports the original report-filter error and failure phase in a failed result; cancellation and disconnected sessions propagate.")]
        PivotReportFilterResult ReadReportFilter(ExcelContext ctx, CancellationToken ct)
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotTables? pivotTables = null;
            Excel.PivotTable? pivot = null;
            Excel.CubeFields? cubeFields = null;
            Excel.CubeField? cubeField = null;
            Excel.PivotFields? pivotFields = null;
            Excel.PivotField? pivotField = null;
            string failurePhase = "resolve-target";

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
                cubeField = GetReportFilterCubeField(cubeFields, fieldName);
                if (Convert.ToInt32(cubeField.Orientation,
                        System.Globalization.CultureInfo.InvariantCulture) != XlPivotFieldOrientation.xlPageField)
                {
                    throw new InvalidOperationException(
                        $"OLAP field '{fieldName}' must be in the report-filter area before its selection can be read.");
                }

                pivotFields = cubeField.PivotFields;
                if (pivotFields.Count == 0)
                {
                    throw new InvalidOperationException(
                        $"OLAP report-filter field '{fieldName}' has no PivotField available.");
                }
                pivotField = pivotFields.Item(1);

                failurePhase = "read-filter-state";
                ReadReportFilterSelection(cubeField, pivotField, result, ct, batch as IExcelBatchWindowControl);
                result.Success = true;
            }
            catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
            {
                result.Success = false;
                result.FailurePhase = failurePhase;
                result.ErrorMessage =
                    $"Failed to read report filter '{fieldName}' on PivotTable '{pivotTableName}' " +
                    $"in worksheet '{sheetName}' during {failurePhase}: {ex.Message}";
            }
            finally
            {
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
        }
    }

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

        return batch.Execute(WriteReportFilter);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "Preserves mutation and update-restoration errors, reports the failed phase and possible partial state, and appends readback failures; cancellation and disconnection propagate.")]
        PivotReportFilterResult WriteReportFilter(ExcelContext ctx, CancellationToken ct)
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
                RequestedItems = [.. selectedItems],
                FilePath = batch.WorkbookPath
            };
            var filterMutationStarted = false;
            var manualUpdateChanged = false;
            var previousManualUpdate = false;
            var previousMultiplePageItems = false;
            string failurePhase = "resolve-target";

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
                cubeField = GetReportFilterCubeField(cubeFields, fieldName);
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

                failurePhase = "capture-state";
                previousManualUpdate = pivot.ManualUpdate;
                previousMultiplePageItems = cubeField.EnableMultiplePageItems;
                failurePhase = "defer-update";
                if (!previousManualUpdate)
                {
                    pivot.ManualUpdate = true;
                    manualUpdateChanged = true;
                }

                Exception? mutationFailure = null;
                string? mutationFailurePhase = null;
                try
                {
                    filterMutationStarted = true;
                    if (selectedItems.Count == 0)
                    {
                        failurePhase = "clear-selection";
                        pivotField.ClearAllFilters();
                    }
                    else if (selectedItems.Count == 1 && !previousMultiplePageItems)
                    {
                        failurePhase = "set-single-selection";
                        pivotField.CurrentPageName = selectedItems[0];
                    }
                    else
                    {
                        if (!previousMultiplePageItems)
                        {
                            failurePhase = "enable-multiple-selection";
                            cubeField.EnableMultiplePageItems = true;
                        }
                        failurePhase = "assign-selected-items";
                        pivotField.VisibleItemsList = selectedItems.ToArray();
                    }
                }
                catch (Exception ex)
                {
                    mutationFailure = ex;
                    mutationFailurePhase = failurePhase;
                }
                if (manualUpdateChanged)
                {
                    failurePhase = "restore-update";
                    OperationCleanup.Run(() => pivot.ManualUpdate = previousManualUpdate, mutationFailure);
                }
                if (mutationFailure is not null)
                {
                    failurePhase = mutationFailurePhase ?? failurePhase;
                    System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(mutationFailure).Throw();
                }

                failurePhase = "read-filter-state";
                ReadReportFilterSelection(cubeField, pivotField, result, ct, batch as IExcelBatchWindowControl);
                result.Success = true;
            }
            catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
            {
                result.Success = false;
                result.FailurePhase = failurePhase;
                result.ErrorMessage =
                    $"Failed to set report filter '{fieldName}' on PivotTable '{pivotTableName}' " +
                    $"in worksheet '{sheetName}' during {failurePhase}: {ex.Message}";
                result.MayHavePartiallyChanged = filterMutationStarted;
                if (filterMutationStarted && cubeField is not null && pivotField is not null)
                {
                    failurePhase = "read-filter-state-after-failure";
                    try
                    {
                        ReadReportFilterSelection(cubeField, pivotField, result, ct, batch as IExcelBatchWindowControl);
                    }
                    catch (Exception readException) when (CreatedObjectFailure.CanReport(readException))
                    {
                        result.ErrorMessage +=
                            $" Actual filter selection could not be read: {readException.Message}";
                    }
                }
            }
            finally
            {
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
        }
    }

    private static PivotReportFilterResult? ValidateReportFilterRequest(
        string workbookPath, string sheetName, string pivotTableName, string fieldName,
        List<string> selectedItems)
    {
        string? error = ValidateReportFilterTarget(sheetName, pivotTableName, fieldName);
        if (error is null && selectedItems is null)
        {
            error = "Selected items must be provided; use an empty list to show all members.";
        }
        else if (error is null && selectedItems!.Any(string.IsNullOrWhiteSpace))
        {
            error = "Selected member names must not be empty.";
        }
        else if (error is null && selectedItems!.Count != selectedItems.Distinct(StringComparer.Ordinal).Count())
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
                RequestedItems = selectedItems is null ? null : [.. selectedItems],
                FilePath = workbookPath
            };
    }

    private static string? ValidateReportFilterTarget(
        string sheetName, string pivotTableName, string fieldName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            return "Worksheet name must not be empty.";
        }
        if (string.IsNullOrWhiteSpace(pivotTableName))
        {
            return "PivotTable name must not be empty.";
        }
        return string.IsNullOrWhiteSpace(fieldName) ? "Field name must not be empty." : null;
    }

    private static Excel.CubeField GetReportFilterCubeField(Excel.CubeFields cubeFields, string fieldName)
    {
        try
        {
            return cubeFields[fieldName];
        }
        catch (System.Runtime.InteropServices.COMException ex) when (
            ex.HResult == unchecked((int)0x8002000B))
        {
            throw new ArgumentException(
                $"Field '{fieldName}' was not found. Use the exact hierarchy unique name returned by " +
                "pivottable_field(list-fields), not a level-qualified name.",
                nameof(fieldName),
                ex);
        }
    }

    private static void ReadReportFilterSelection(
        Excel.CubeField cubeField, Excel.PivotField pivotField,
        PivotReportFilterResult result, CancellationToken cancellationToken,
        IExcelBatchWindowControl? windowControl)
    {
        result.FieldArea = PivotFieldArea.Filter;
        result.MultipleSelectionEnabled = cubeField.EnableMultiplePageItems;
        if (!result.MultipleSelectionEnabled)
        {
            string currentPageName = pivotField.CurrentPageName;
            result.ShowAll = PivotTableHelpers.IsAllOlapReportFilterMember(
                cubeField, currentPageName, cancellationToken, windowControl);
            result.SelectedItems = result.ShowAll ? [] : [currentPageName];
            result.SelectionStateAvailable = true;
            return;
        }

        result.ShowAll = cubeField.AllItemsVisible;
        if (result.ShowAll)
        {
            result.SelectedItems.Clear();
            result.SelectionStateAvailable = true;
            return;
        }

        object? visibleItems = pivotField.VisibleItemsList;
        result.SelectedItems = visibleItems is Array items
            ? items.Cast<object>().Select(item => item.ToString() ?? string.Empty).ToList()
            : [];
        result.SelectionStateAvailable = true;
    }

    /// <summary>
    /// Sorts a field
    /// </summary>
    public PivotFieldResult SortField(IExcelBatch batch, string pivotTableName,
        string fieldName, SortDirection direction = SortDirection.Ascending, string? dataFieldName = null)
    {
        if (!Enum.IsDefined(direction))
            throw new ArgumentOutOfRangeException(nameof(direction));
        if (dataFieldName is null)
            return ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
                (strategy, pivot) => strategy.SortField(pivot, fieldName, direction, batch.WorkbookPath));
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        ArgumentException.ThrowIfNullOrWhiteSpace(fieldName);
        ArgumentException.ThrowIfNullOrWhiteSpace(dataFieldName);
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.Worksheet? sheet = null;
            Excel.PivotCache? cache = null;
            Excel.CubeFields? cubes = null;
            Excel.CubeField? cube = null;
            Excel.PivotFields? levels = null;
            Excel.PivotFields? values = null;
            Excel.PivotField? value = null;
            Excel.PivotField? field = null;
            try
            {
                pivot = (Excel.PivotTable)FindPivotTable(ctx.Book, pivotTableName);
                sheet = (Excel.Worksheet)pivot.Parent;
                if (sheet.ProtectContents)
                    throw new InvalidOperationException("Unprotect the PivotTable worksheet before sorting.");
                cache = pivot.PivotCache();
                bool isOlap = cache.OLAP;
                if (isOlap)
                {
                    cubes = pivot.CubeFields;
                    cube = cubes[fieldName];
                    if (cube.Orientation is not (Excel.XlPivotFieldOrientation.xlRowField or Excel.XlPivotFieldOrientation.xlColumnField))
                        throw new ArgumentException("Place the OLAP hierarchy on the row or column axis before sorting.");
                    levels = cube.PivotFields;
                    if (levels.Count == 0)
                        throw new ArgumentException("Place the OLAP hierarchy on the row or column axis before sorting.");
                    field = levels.Item(1);
                }
                else
                {
                    field = FindCalculationBaseField(pivot, fieldName, ct);
                }
                RequireAxisField(field);
                values = (Excel.PivotFields)pivot.DataFields;
                value = FindValueField(values, dataFieldName, ct, allowOlapNames: isOlap);
                string sortBy = isOlap ? value.SourceName : value.Name;
                ct.ThrowIfCancellationRequested();
                var order = direction == SortDirection.Ascending
                    ? Excel.XlSortOrder.xlAscending : Excel.XlSortOrder.xlDescending;
                field.AutoSort((int)order, sortBy);
                if (field.AutoSortOrder != (int)order || field.AutoSortField != sortBy)
                    throw new InvalidOperationException("Excel did not apply the requested Values sort. Inspect get-data before retrying.");
                return new PivotFieldResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    FieldName = fieldName,
                    CustomName = field.Caption,
                    Area = (PivotFieldArea)field.Orientation
                };
            }
            finally
            {
                ComUtilities.Release(ref value);
                ComUtilities.Release(ref values);
                ComUtilities.Release(ref field);
                ComUtilities.Release(ref levels);
                ComUtilities.Release(ref cube);
                ComUtilities.Release(ref cubes);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref pivot);
            }
        });
    }
}
