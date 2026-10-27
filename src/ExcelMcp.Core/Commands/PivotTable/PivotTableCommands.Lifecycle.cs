using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable lifecycle operations (List, GetInfo, Create, Delete, Refresh)
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Lists all PivotTables in workbook
    /// </summary>
    public PivotTableListResult List(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            var pivotTables = new List<PivotTableInfo>();
            Excel.Sheets? sheets = null;

            try
            {
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Worksheet? sheet = null;
                    Excel.PivotTables? pivotTablesCol = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[i];
                        string sheetName = sheet.Name;
                        pivotTablesCol = (Excel.PivotTables)sheet.PivotTables();

                        for (int j = 1; j <= pivotTablesCol.Count; j++)
                        {
                            ct.ThrowIfCancellationRequested();
                            Excel.PivotTable? pivot = null;
                            Excel.PivotCache? pivotCache = null;
                            Excel.Range? tableRange = null;
                            Excel.PivotFields? rowFields = null;
                            Excel.PivotFields? columnFields = null;
                            Excel.PivotFields? dataFields = null;
                            Excel.PivotFields? pageFields = null;
                            try
                            {
                                pivot = pivotTablesCol.Item(j);

                                // Get basic info (should never fail)
                                string pivotName = pivot.Name;
                                int cacheIndex = Convert.ToInt32(pivot.CacheIndex);

                                // Get properties with defensive error handling
                                string? range = null;
                                string? layoutError = null;
                                string? sourceData = null;
                                int rowFieldCount = 0;
                                int columnFieldCount = 0;
                                int valueFieldCount = 0;
                                int filterFieldCount = 0;
                                DateTime? lastRefresh = null;

                                tableRange = pivot.TableRange2;
                                range = tableRange.Address;
                                layoutError = PivotTableHelpers.GetLayoutError((Microsoft.Office.Interop.Excel.PivotTable)pivot);
                                pivotCache = pivot.PivotCache();
                                if ((int)pivotCache.SourceType == (int)Excel.XlPivotTableSourceType.xlDatabase)
                                    sourceData = ((object?)pivotCache.SourceData)?.ToString() ?? string.Empty;
                                else
                                    sourceData = "(external or Data Model)";
                                lastRefresh = GetRefreshDateSafe(pivotCache.RefreshDate);
                                rowFields = (Excel.PivotFields)pivot.RowFields;
                                columnFields = (Excel.PivotFields)pivot.ColumnFields;
                                dataFields = (Excel.PivotFields)pivot.DataFields;
                                pageFields = (Excel.PivotFields)pivot.PageFields;
                                rowFieldCount = Convert.ToInt32(rowFields.Count);
                                columnFieldCount = Convert.ToInt32(columnFields.Count);
                                valueFieldCount = Convert.ToInt32(dataFields.Count);
                                filterFieldCount = Convert.ToInt32(pageFields.Count);

                                var info = new PivotTableInfo
                                {
                                    Name = pivotName,
                                    CacheIndex = cacheIndex,
                                    SheetName = sheetName,
                                    Range = range ?? "(unavailable)",
                                    LayoutError = layoutError,
                                    SourceData = sourceData ?? string.Empty,
                                    RowFieldCount = rowFieldCount,
                                    ColumnFieldCount = columnFieldCount,
                                    ValueFieldCount = valueFieldCount,
                                    FilterFieldCount = filterFieldCount,
                                    LastRefresh = lastRefresh
                                };

                                pivotTables.Add(info);
                            }
                            finally
                            {
                                ComUtilities.Release(ref pageFields);
                                ComUtilities.Release(ref dataFields);
                                ComUtilities.Release(ref columnFields);
                                ComUtilities.Release(ref rowFields);
                                ComUtilities.Release(ref tableRange);
                                ComUtilities.Release(ref pivotCache);
                                ComUtilities.Release(ref pivot);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref pivotTablesCol);
                        ComUtilities.Release(ref sheet);
                    }
                }

                return new PivotTableListResult
                {
                    Success = true,
                    PivotTables = pivotTables,
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <summary>
    /// Gets detailed information about a PivotTable
    /// </summary>
    public PivotTableInfoResult Read(
        IExcelBatch batch,
        string pivotTableName,
        string? sheetName = null,
        int maxSelectedMembers = PivotTableSnapshotHelpers.DefaultMaxSelectedMembers)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(pivotTableName);
        if (sheetName is not null)
            ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        PivotTableSnapshotHelpers.ValidateMaxMembers(maxSelectedMembers);

        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? pivotCache = null;
            Excel.CubeFields? cubeFields = null;
            Excel.PivotFields? pivotFields = null;
            Excel.Worksheet? parentSheet = null;
            Excel.Range? tableRange = null;
            Excel.PivotFields? rowFields = null;
            Excel.PivotFields? columnFields = null;
            Excel.PivotFields? dataFields = null;
            Excel.PivotFields? pageFields = null;

            try
            {
                pivot = sheetName is null
                    ? FindPivotTable(ctx.Book, pivotTableName)
                    : FindPivotTable(ctx.Book, sheetName, pivotTableName);
                pivotCache = pivot.PivotCache();
                parentSheet = (Excel.Worksheet)pivot.Parent;
                tableRange = pivot.TableRange2;
                rowFields = (Excel.PivotFields)pivot.RowFields;
                columnFields = (Excel.PivotFields)pivot.ColumnFields;
                dataFields = (Excel.PivotFields)pivot.DataFields;
                pageFields = (Excel.PivotFields)pivot.PageFields;

                // Get basic info with defensive error handling (properties can throw on Data Model sources)
                var info = new PivotTableInfo
                {
                    Name = pivot.Name,
                    CacheIndex = Convert.ToInt32(pivot.CacheIndex),
                    SheetName = parentSheet.Name,
                    Range = tableRange.Address,
                    LayoutError = PivotTableHelpers.GetLayoutError((Excel.PivotTable)pivot),
                    RowFieldCount = Convert.ToInt32(rowFields.Count),
                    ColumnFieldCount = Convert.ToInt32(columnFields.Count),
                    ValueFieldCount = Convert.ToInt32(dataFields.Count),
                    FilterFieldCount = Convert.ToInt32(pageFields.Count),
                    LastRefresh = GetRefreshDateSafe(pivotCache.RefreshDate)
                };

                if ((int)pivotCache.SourceType == (int)Excel.XlPivotTableSourceType.xlDatabase)
                    info.SourceData = ((object?)pivotCache.SourceData)?.ToString() ?? string.Empty;
                else
                    info.SourceData = "[Data Model Source]";

                var olapState = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
                    () => Convert.ToBoolean(pivotCache.OLAP));
                bool isOlap = olapState.Available && olapState.Value;

                var fieldSnapshot = olapState.Available
                    ? PivotTableSnapshotHelpers.TryReadSnapshotProperty(() =>
                    {
                        List<PivotFieldInfo> fields;
                        if (isOlap)
                        {
                            cubeFields = (Excel.CubeFields)pivot.CubeFields;
                            fields = GetCubeFieldsInfo(cubeFields);
                        }
                        else
                        {
                            pivotFields = (Excel.PivotFields)pivot.PivotFields();
                            fields = GetRegularFieldsInfo(pivotFields);
                        }

                        var selections = GetFieldSelectionSnapshots(
                            isOlap,
                            (Excel.CubeFields?)cubeFields,
                            (Excel.PivotFields?)pivotFields,
                            pivotTableName,
                            batch.WorkbookPath,
                            maxSelectedMembers,
                            ct,
                            batch as IExcelBatchWindowControl);
                        return (Fields: fields, Selections: selections);
                    })
                    : new PivotSnapshotProperty<(List<PivotFieldInfo> Fields, List<PivotFieldSelectionSnapshot> Selections)>(
                        false, default, olapState.ErrorMessage);
                var fieldsResult = fieldSnapshot.Value;
                var valueFieldsState = PivotTableSnapshotHelpers.TryReadSnapshotProperty(
                    () => ReadAllValueFieldCalculations((Excel.PivotTable)pivot, batch.WorkbookPath, ct));

                var fields = fieldSnapshot.Available ? fieldsResult.Fields : [];
                var fieldSelections = fieldSnapshot.Available ? fieldsResult.Selections : [];
                return new PivotTableInfoResult
                {
                    Success = true,
                    PivotTable = info,
                    Fields = fields,
                    FilePath = batch.WorkbookPath,
                    FieldSelections = fieldSelections,
                    ValueFields = valueFieldsState.Value ?? [],
                    IsOlap = isOlap,
                    IsOlapAvailable = olapState.Available,
                    IsOlapError = olapState.ErrorMessage,
                    FieldSelectionsAvailable = fieldSnapshot.Available,
                    FieldSelectionsError = fieldSnapshot.ErrorMessage,
                    ValueFieldsAvailable = valueFieldsState.Available,
                    ValueFieldsError = valueFieldsState.ErrorMessage,
                    MaxSelectedMembers = maxSelectedMembers,
                    SnapshotComplete = olapState.Available &&
                        fieldSnapshot.Available &&
                        valueFieldsState.Available &&
                        fieldSelections.All(selection =>
                            selection.FieldIdentityAvailable &&
                            selection.LayoutAvailable &&
                            selection.SelectionStateAvailable &&
                            selection.FiltersAvailable &&
                            !selection.MembersTruncated)
                };
            }
            finally
            {
                ComUtilities.Release(ref pageFields);
                ComUtilities.Release(ref dataFields);
                ComUtilities.Release(ref columnFields);
                ComUtilities.Release(ref rowFields);
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref parentSheet);
                ComUtilities.Release(ref cubeFields);
                ComUtilities.Release(ref pivotFields);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Gets field info from CubeFields (OLAP/Data Model PivotTables)
    /// </summary>
    private static List<PivotFieldInfo> GetCubeFieldsInfo(Excel.CubeFields cubeFields)
    {
        var fields = new List<PivotFieldInfo>();
        int fieldCount = cubeFields.Count;
        for (int i = 1; i <= fieldCount; i++)
        {
            Microsoft.Office.Interop.Excel.CubeField? cubeField = null;
            try
            {
                cubeField = cubeFields[i];
                var orientation = cubeField.Orientation;
                fields.Add(new PivotFieldInfo
                {
                    Name = cubeField.Name,
                    Area = (PivotFieldArea)orientation,
                    CustomName = cubeField.Caption,
                    Position = orientation == Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlHidden
                        ? 0 : cubeField.Position,
                    DataType = "Cube"
                });
            }
            finally
            {
                ComUtilities.Release(ref cubeField);
            }
        }
        return fields;
    }

    /// <summary>
    /// Gets field info from PivotFields (regular PivotTables)
    /// </summary>
    private static List<PivotFieldInfo> GetRegularFieldsInfo(Excel.PivotFields pivotFields)
    {
        var fields = new List<PivotFieldInfo>();

        int fieldCount = pivotFields.Count;

        for (int i = 1; i <= fieldCount; i++)
        {
            Excel.PivotField? field = null;
            try
            {
                field = pivotFields.Item(i);
                object? sourceName = field.SourceName;
                if (PivotTableHelpers.IsValuesPlaceholderField(field, sourceName))
                    continue; // Excel's synthetic Values field is not a source column.
                int orientation = (int)field.Orientation;

                var fieldInfo = new PivotFieldInfo
                {
                    Name = sourceName?.ToString() ?? field.Name?.ToString() ?? $"Field{i}",
                    CustomName = field.Caption?.ToString() ?? string.Empty,
                    Area = orientation switch
                    {
                        XlPivotFieldOrientation.xlRowField => PivotFieldArea.Row,
                        XlPivotFieldOrientation.xlColumnField => PivotFieldArea.Column,
                        XlPivotFieldOrientation.xlPageField => PivotFieldArea.Filter,
                        XlPivotFieldOrientation.xlDataField => PivotFieldArea.Value,
                        _ => PivotFieldArea.Hidden
                    },
                    Position = orientation != XlPivotFieldOrientation.xlHidden ? Convert.ToInt32((object)field.Position, System.Globalization.CultureInfo.InvariantCulture) : 0,
                    DataType = PivotTableHelpers.DetectFieldDataType(field)
                };

                // Get function for value fields
                if (orientation == XlPivotFieldOrientation.xlDataField)
                {
                    int comFunction = (int)field.Function;
                    fieldInfo.Function = PivotTableHelpers.GetAggregationFunctionFromCom(comFunction);
                }

                fields.Add(fieldInfo);
            }
            finally
            {
                ComUtilities.Release(ref field);
            }
        }

        return fields;
    }

    /// <summary>
    /// Deletes a PivotTable
    /// </summary>
    public OperationResult Delete(IExcelBatch batch, string pivotTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.Range? tableRange = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                tableRange = pivot.TableRange2;

                // Delete the PivotTable
                tableRange.Clear();

                return new OperationResult
                {
                    Success = true,
                    Action = "delete",
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
    /// Refreshes a PivotTable
    /// </summary>
    public PivotTableRefreshResult Refresh(IExcelBatch batch, string pivotTableName, TimeSpan? timeout = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotCache? pivotCache = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                pivotCache = pivot.PivotCache();

                int previousRecordCount = pivotCache.RecordCount;

                // Refresh the PivotTable
                pivot.RefreshTable();

                int currentRecordCount = pivotCache.RecordCount;

                return new PivotTableRefreshResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    RefreshTime = DateTime.Now,
                    SourceRecordCount = currentRecordCount,
                    PreviousRecordCount = previousRecordCount,
                    StructureChanged = currentRecordCount != previousRecordCount,
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Safely converts Excel RefreshDate (which can be DateTime or double OLE date) to DateTime?
    /// </summary>
    private static DateTime? GetRefreshDateSafe(object? refreshDate)
    {
        if (refreshDate == null)
            return null;

        if (refreshDate is DateTime dt)
            return dt;

        if (refreshDate is double dbl)
            return DateTime.FromOADate(dbl);

        return null;
    }
}
