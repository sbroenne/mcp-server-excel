using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable lifecycle operations (List, GetInfo, Create, Delete, Refresh)
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Lists all PivotTables in workbook
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotTableListResult> ListAsync(IExcelBatch batch)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            var pivotTables = new List<PivotTableInfo>();
            dynamic? sheets = null;

            try
            {
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    dynamic? pivotTablesCol = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        string sheetName = sheet.Name;
                        pivotTablesCol = sheet.PivotTables;

                        for (int j = 1; j <= pivotTablesCol.Count; j++)
                        {
                            dynamic? pivot = null;
                            dynamic? pivotCache = null;
                            try
                            {
                                pivot = pivotTablesCol.Item(j);
                                pivotCache = pivot.PivotCache;

                                // Handle RefreshDate which can be DateTime or double (OLE date)
                                DateTime? lastRefresh = null;
                                if (pivotCache.RefreshDate != null)
                                {
                                    var refreshDate = pivotCache.RefreshDate;
                                    if (refreshDate is DateTime dt)
                                    {
                                        lastRefresh = dt;
                                    }
                                    else if (refreshDate is double dbl)
                                    {
                                        lastRefresh = DateTime.FromOADate(dbl);
                                    }
                                }

                                var info = new PivotTableInfo
                                {
                                    Name = pivot.Name,
                                    SheetName = sheetName,
                                    Range = pivot.TableRange2.Address,
                                    SourceData = pivotCache.SourceData?.ToString() ?? string.Empty,
                                    RowFieldCount = pivot.RowFields.Count,
                                    ColumnFieldCount = pivot.ColumnFields.Count,
                                    ValueFieldCount = pivot.DataFields.Count,
                                    FilterFieldCount = pivot.PageFields.Count,
                                    LastRefresh = lastRefresh
                                };

                                pivotTables.Add(info);
                            }
                            finally
                            {
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
            catch (Exception ex)
            {
                return new PivotTableListResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to list PivotTables: {ex.Message}",
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
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotTableInfoResult> GetInfoAsync(IExcelBatch batch, string pivotTableName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? pivotCache = null;
            dynamic? pivotFields = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                pivotCache = pivot.PivotCache;
                pivotFields = pivot.PivotFields;

                // Get basic info
                var info = new PivotTableInfo
                {
                    Name = pivot.Name,
                    SheetName = pivot.Parent.Name,
                    Range = pivot.TableRange2.Address,
                    SourceData = pivotCache.SourceData?.ToString() ?? string.Empty,
                    RowFieldCount = pivot.RowFields.Count,
                    ColumnFieldCount = pivot.ColumnFields.Count,
                    ValueFieldCount = pivot.DataFields.Count,
                    FilterFieldCount = pivot.PageFields.Count,
                    LastRefresh = GetRefreshDateSafe(pivotCache.RefreshDate)
                };

                // Get field details
                var fields = new List<PivotFieldInfo>();
                for (int i = 1; i <= pivotFields.Count; i++)
                {
                    dynamic? field = null;
                    try
                    {
                        field = pivotFields.Item(i);
                        int orientation = Convert.ToInt32(field.Orientation);

                        var fieldInfo = new PivotFieldInfo
                        {
                            Name = field.SourceName?.ToString() ?? field.Name?.ToString() ?? $"Field{i}",
                            CustomName = field.Caption?.ToString() ?? string.Empty,
                            Area = orientation switch
                            {
                                XlPivotFieldOrientation.xlRowField => PivotFieldArea.Row,
                                XlPivotFieldOrientation.xlColumnField => PivotFieldArea.Column,
                                XlPivotFieldOrientation.xlPageField => PivotFieldArea.Filter,
                                XlPivotFieldOrientation.xlDataField => PivotFieldArea.Value,
                                _ => PivotFieldArea.Hidden
                            },
                            Position = orientation != XlPivotFieldOrientation.xlHidden ? Convert.ToInt32(field.Position) : 0,
                            DataType = DetectFieldDataType(field)
                        };

                        // Get function for value fields
                        if (orientation == XlPivotFieldOrientation.xlDataField)
                        {
                            int comFunction = Convert.ToInt32(field.Function);
                            fieldInfo.Function = GetAggregationFunctionFromCom(comFunction);
                        }

                        fields.Add(fieldInfo);
                    }
                    finally
                    {
                        ComUtilities.Release(ref field);
                    }
                }

                return new PivotTableInfoResult
                {
                    Success = true,
                    PivotTable = info,
                    Fields = fields,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                return new PivotTableInfoResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to get PivotTable info: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivotFields);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Deletes a PivotTable
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string pivotTableName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? tableRange = null;

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
            catch (Exception ex)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to delete PivotTable: {ex.Message}",
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
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotTableRefreshResult> RefreshAsync(IExcelBatch batch, string pivotTableName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? pivot = null;
            dynamic? pivotCache = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                pivotCache = pivot.PivotCache;

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
            catch (Exception ex)
            {
                return new PivotTableRefreshResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to refresh PivotTable: {ex.Message}",
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
    private static DateTime? GetRefreshDateSafe(dynamic refreshDate)
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

