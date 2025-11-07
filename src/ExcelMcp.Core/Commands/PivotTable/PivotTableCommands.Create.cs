using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable creation operations
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Creates a PivotTable from an Excel range
    /// Following VBA pattern from ReneNyffenegger/about-MS-Office-object-model
    /// </summary>
    public async Task<PivotTableCreateResult> CreateFromRangeAsync(IExcelBatch batch,
        string sourceSheet, string sourceRange,
        string destinationSheet, string destinationCell,
        string pivotTableName)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? sourceWorksheet = null;
            dynamic? sourceRangeObj = null;
            dynamic? destWorksheet = null;
            dynamic? destRangeObj = null;
            dynamic? pivotCaches = null;
            dynamic? pivotCache = null;
            dynamic? pivotTable = null;

            try
            {
                // STEP 1: Validate source range has headers and data
                sourceWorksheet = ctx.Book.Worksheets.Item(sourceSheet);
                sourceRangeObj = sourceWorksheet.Range[sourceRange];

                if (sourceRangeObj.Rows.Count < 2)
                {
                    throw new InvalidOperationException($"Source range must contain headers and at least one data row. Found {sourceRangeObj.Rows.Count} rows");
                }

                // STEP 2: Create PivotCache from source range
                // VBA: Set pivot_cache = activeWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="csv_data", Version:=xlPivotTableVersion14)
                pivotCaches = ctx.Book.PivotCaches();
                string sourceDataRef = $"{sourceSheet}!{sourceRange}";

                // xlDatabase = 1, xlPivotTableVersion14 = 4
                pivotCache = pivotCaches.Create(
                    SourceType: 1,
                    SourceData: sourceDataRef,
                    Version: 4
                );

                // STEP 3: Create PivotTable from cache
                // VBA: Set pivot_table = pivot_cache.CreatePivotTable(TableDestination:=pivot_table_upper_left)
                destWorksheet = ctx.Book.Worksheets.Item(destinationSheet);
                destRangeObj = destWorksheet.Range[destinationCell];

                pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: destRangeObj,
                    TableName: pivotTableName
                );

                // STEP 4: Refresh to materialize the PivotTable structure
                pivotTable.RefreshTable();

                // STEP 5: Get available fields from PivotTable (VBA pattern)
                // VBA: Set pf_col_1 = pivot_table.PivotFields("col_1")
                // STEP 5: Get available fields from source range headers
                // These are the fields that CAN be added to the PivotTable
                var availableFields = new List<string>();

                dynamic? headerRow = null;
                try
                {
                    headerRow = sourceRangeObj.Rows[1];
                    object[,] headers = headerRow.Value2;

                    for (int col = 1; col <= headers.GetLength(1); col++)
                    {
                        var header = headers[1, col]?.ToString();
                        if (!string.IsNullOrWhiteSpace(header))
                        {
                            availableFields.Add(header);
                        }
                    }

                    if (availableFields.Count == 0)
                    {
                        throw new InvalidOperationException($"No field headers found in source range. Header row has {headers.GetLength(1)} columns.");
                    }
                }
                finally
                {
                    ComUtilities.Release(ref headerRow);
                }

                return new PivotTableCreateResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    SheetName = destinationSheet,
                    Range = pivotTable.TableRange2.Address,
                    SourceData = sourceDataRef,
                    SourceRowCount = sourceRangeObj.Rows.Count - 1,
                    AvailableFields = availableFields,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                // Cleanup on failure
                if (pivotTable != null)
                {
                    try
                    {
                        dynamic? tableRange = null;
                        try
                        {
                            tableRange = pivotTable.TableRange2;
                            tableRange.Clear();
                        }
                        finally
                        {
                            ComUtilities.Release(ref tableRange);
                        }
                    }
                    catch { /* Ignore cleanup errors */ }
                }

                return new PivotTableCreateResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to create PivotTable: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
                ComUtilities.Release(ref destRangeObj);
                ComUtilities.Release(ref destWorksheet);
                ComUtilities.Release(ref sourceRangeObj);
                ComUtilities.Release(ref sourceWorksheet);
            }
        });
    }

    /// <summary>
    /// Creates a PivotTable from an Excel Table
    /// </summary>
    public async Task<PivotTableCreateResult> CreateFromTableAsync(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? destWorksheet = null;
            dynamic? destRangeObj = null;
            dynamic? pivotCaches = null;
            dynamic? pivotCache = null;
            dynamic? pivotTable = null;

            try
            {
                // Find the table
                dynamic? sheets = null;
                bool tableFound = false;

                try
                {
                    sheets = ctx.Book.Worksheets;
                    for (int i = 1; i <= sheets.Count; i++)
                    {
                        dynamic? sheet = null;
                        dynamic? listObjects = null;
                        try
                        {
                            sheet = sheets.Item(i);
                            listObjects = sheet.ListObjects;

                            for (int j = 1; j <= listObjects.Count; j++)
                            {
                                dynamic? tbl = null;
                                try
                                {
                                    tbl = listObjects.Item(j);
                                    if (tbl.Name == tableName)
                                    {
                                        table = tbl;
                                        tableFound = true;
                                        break;
                                    }
                                }
                                finally
                                {
                                    if (tbl != null && tbl != table)
                                    {
                                        ComUtilities.Release(ref tbl);
                                    }
                                }
                            }

                            if (tableFound) break;
                        }
                        finally
                        {
                            ComUtilities.Release(ref listObjects);
                            ComUtilities.Release(ref sheet);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref sheets);
                }

                if (!tableFound || table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found in workbook");
                }

                // Get table range and headers
                dynamic? tableRange = null;
                dynamic? headerRow = null;
                var headers = new List<string>();
                int rowCount = 0;

                try
                {
                    tableRange = table.Range;
                    rowCount = tableRange.Rows.Count;

                    if (rowCount < 2)
                    {
                        throw new InvalidOperationException($"Table '{tableName}' must contain at least one data row (has {rowCount} rows including header)");
                    }

                    // Get headers
                    dynamic? headerRowCol = null;
                    try
                    {
                        headerRowCol = table.HeaderRowRange;
                        object[,] headerValues = headerRowCol.Value2;

                        for (int col = 1; col <= headerValues.GetLength(1); col++)
                        {
                            var header = headerValues[1, col]?.ToString();
                            if (!string.IsNullOrWhiteSpace(header))
                            {
                                headers.Add(header);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref headerRowCol);
                    }

                    // Create PivotCache from table
                    pivotCaches = ctx.Book.PivotCaches();
                    string sourceDataRef = $"{table.Parent.Name}!{table.Name}";

                    // xlDatabase = 1
                    pivotCache = pivotCaches.Create(
                        SourceType: 1,
                        SourceData: sourceDataRef
                    );

                    // Create PivotTable
                    destWorksheet = ctx.Book.Worksheets.Item(destinationSheet);
                    destRangeObj = destWorksheet.Range[destinationCell];

                    pivotTable = pivotCache.CreatePivotTable(
                        TableDestination: destRangeObj,
                        TableName: pivotTableName
                    );

                    // Refresh to materialize layout
                    pivotTable.RefreshTable();

                    return new PivotTableCreateResult
                    {
                        Success = true,
                        PivotTableName = pivotTableName,
                        SheetName = destinationSheet,
                        Range = pivotTable.TableRange2.Address,
                        SourceData = sourceDataRef,
                        SourceRowCount = rowCount - 1, // Exclude header
                        AvailableFields = headers,
                        FilePath = batch.WorkbookPath
                    };
                }
                finally
                {
                    ComUtilities.Release(ref headerRow);
                    ComUtilities.Release(ref tableRange);
                }
            }
            catch (Exception ex)
            {
                return new PivotTableCreateResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to create PivotTable from table: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
                ComUtilities.Release(ref destRangeObj);
                ComUtilities.Release(ref destWorksheet);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Creates a PivotTable from a Power Pivot Data Model table
    /// Uses xlExternal source type with "ThisWorkbookDataModel" connection
    /// </summary>
    public async Task<PivotTableCreateResult> CreateFromDataModelAsync(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName)
    {
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? modelTable = null;
            dynamic? destWorksheet = null;
            dynamic? destRangeObj = null;
            dynamic? pivotCaches = null;
            dynamic? pivotCache = null;
            dynamic? pivotTable = null;

            try
            {
                // STEP 1: Verify Data Model exists and find the table
                // NOTE: Every workbook has a Model object, but it may be empty
                model = ctx.Book.Model;

                // Find the table in the Data Model
                dynamic? modelTables = null;
                bool tableFound = false;
                try
                {
                    modelTables = model.ModelTables;

                    // Check if Data Model has any tables
                    if (modelTables == null || modelTables.Count == 0)
                    {
                        throw new InvalidOperationException("Workbook does not contain a Power Pivot Data Model");
                    }

                    for (int i = 1; i <= modelTables.Count; i++)
                    {
                        dynamic? tbl = null;
                        try
                        {
                            tbl = modelTables.Item(i);
                            if (tbl.Name == tableName)
                            {
                                modelTable = tbl;
                                tableFound = true;
                                break;
                            }
                        }
                        finally
                        {
                            if (tbl != null && tbl != modelTable)
                            {
                                ComUtilities.Release(ref tbl);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                }

                if (!tableFound || modelTable == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found in Data Model");
                }

                // Get columns from the Data Model table
                var headers = new List<string>();
                int recordCount = 0;

                try
                {
                    recordCount = ComUtilities.SafeGetInt(modelTable, "RecordCount");

                    // Get columns
                    dynamic? modelColumns = null;
                    try
                    {
                        modelColumns = modelTable.ModelTableColumns;
                        for (int i = 1; i <= modelColumns.Count; i++)
                        {
                            dynamic? column = null;
                            try
                            {
                                column = modelColumns.Item(i);
                                var colName = ComUtilities.SafeGetString(column, "Name");
                                if (!string.IsNullOrWhiteSpace(colName))
                                {
                                    headers.Add(colName);
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref column);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref modelColumns);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to read columns from Data Model table '{tableName}': {ex.Message}");
                }

                if (headers.Count == 0)
                {
                    throw new InvalidOperationException($"Data Model table '{tableName}' has no columns");
                }

                // STEP 2: Create PivotCache from Data Model
                // Using xlExternal (2) with "ThisWorkbookDataModel" connection
                pivotCaches = ctx.Book.PivotCaches();

                // xlExternal = 2
                pivotCache = pivotCaches.Create(
                    SourceType: 2,
                    SourceData: "ThisWorkbookDataModel"
                );

                // STEP 3: Create PivotTable from cache
                destWorksheet = ctx.Book.Worksheets.Item(destinationSheet);
                destRangeObj = destWorksheet.Range[destinationCell];

                pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: destRangeObj,
                    TableName: pivotTableName
                );

                // STEP 4: Refresh to materialize the PivotTable structure
                pivotTable.RefreshTable();

                return new PivotTableCreateResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    SheetName = destinationSheet,
                    Range = pivotTable.TableRange2.Address,
                    SourceData = $"ThisWorkbookDataModel[{tableName}]",
                    SourceRowCount = recordCount,
                    AvailableFields = headers,
                    FilePath = batch.WorkbookPath
                };
            }
            catch (Exception ex)
            {
                // Cleanup on failure
                if (pivotTable != null)
                {
                    try
                    {
                        dynamic? tableRange = null;
                        try
                        {
                            tableRange = pivotTable.TableRange2;
                            tableRange.Clear();
                        }
                        finally
                        {
                            ComUtilities.Release(ref tableRange);
                        }
                    }
                    catch { /* Ignore cleanup errors */ }
                }

                return new PivotTableCreateResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to create PivotTable from Data Model table: {ex.Message}",
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
                ComUtilities.Release(ref destRangeObj);
                ComUtilities.Release(ref destWorksheet);
                ComUtilities.Release(ref modelTable);
                ComUtilities.Release(ref model);
            }
        });
    }
}
