using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

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
    public PivotTableCreateResult CreateFromRange(IExcelBatch batch,
        string sourceSheet, string sourceRange,
        string destinationSheet, string destinationCell,
        string pivotTableName, string? shareCacheWith = null, string? shareCacheWithSheet = null)
    {
        ValidateCacheSharingInputs(shareCacheWith, shareCacheWithSheet);
        if (shareCacheWith is not null && !IsFixedA1SourceAddress(sourceRange))
            throw new ArgumentException(
                "Cache sharing from a range requires an explicit fixed sourceRange address. Named and structured source definitions cannot be substituted for fixed ranges; use create-from-table for an Excel Table.");
        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(5));

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sourceWorksheet = null;
            Excel.Range? sourceRangeObj = null;
            Excel.Worksheet? destWorksheet = null;
            Excel.Range? destRangeObj = null;
            Excel.PivotCaches? pivotCaches = null;
            Excel.PivotCache? pivotCache = null;
            Excel.PivotTable? pivotTable = null;
            Excel.Range? sourceRows = null;

            try
            {
                sourceWorksheet = Utilities.WorksheetLookup.Get(ctx.Book, sourceSheet);
                sourceRangeObj = sourceWorksheet.Range[sourceRange];
                sourceRows = sourceRangeObj.Rows;
                int sourceRowCount = sourceRows.Count;

                if (sourceRowCount < 2)
                {
                    throw new InvalidOperationException($"Source range must contain headers and at least one data row. Found {sourceRowCount} rows");
                }

                var availableFields = new List<string>();
                var blankHeaderColumns = new List<int>();

                Excel.Range? headerRow = null;
                try
                {
                    headerRow = (Excel.Range)sourceRows[1];
                    var headerValues = ExcelValueNormalizer.Normalize((object?)headerRow.Value2);

                    for (int column = 0; column < headerValues.ColumnCount; column++)
                    {
                        var header = headerValues.Values[0][column]?.ToString();
                        if (string.IsNullOrWhiteSpace(header))
                        {
                            blankHeaderColumns.Add(column + 1);
                        }
                        else
                        {
                            availableFields.Add(header);
                        }
                    }

                    if (availableFields.Count == 0)
                    {
                        throw new InvalidOperationException($"No field headers found in source range. Header row has {headerValues.ColumnCount} columns.");
                    }

                    if (blankHeaderColumns.Count > 0)
                    {
                        throw new InvalidOperationException(
                            $"Every source column must have a field header. Header row contains blank field name(s) in column(s): " +
                            $"{string.Join(", ", blankHeaderColumns)}.");
                    }
                }
                finally
                {
                    ComUtilities.Release(ref headerRow);
                }

                destWorksheet = Utilities.WorksheetLookup.Get(ctx.Book, destinationSheet);
                EnsurePivotTableNameAvailable(destWorksheet, pivotTableName, destinationSheet);
                destRangeObj = destWorksheet.Range[destinationCell];

                string sourceDataRef = $"'{sourceSheet.Replace("'", "''", StringComparison.Ordinal)}'!{sourceRange}";

                if (shareCacheWith is not null)
                    pivotCache = ResolveSharedCreationCache(ctx.Book, ctx.App, shareCacheWithSheet!, shareCacheWith,
                        (Excel.Worksheet)sourceWorksheet, (Excel.Range)sourceRangeObj, null);
                else
                {
                    pivotCaches = ctx.Book.PivotCaches();
                    pivotCache = (Excel.PivotCache)pivotCaches.Create(
                        SourceType: Excel.XlPivotTableSourceType.xlDatabase,
                        SourceData: sourceDataRef,
                        Version: 4
                    );
                }

                pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: destRangeObj,
                    TableName: pivotTableName
                );

                string rangeAddress = FinishCreatedPivotTable(pivotTable, pivotTableName, destinationSheet);
                return new PivotTableCreateResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    SheetName = destinationSheet,
                    Range = rangeAddress,
                    SourceData = sourceDataRef,
                    SourceRowCount = sourceRowCount - 1,
                    AvailableFields = availableFields,
                    FilePath = batch.WorkbookPath
                };
            }
            finally
            {
                ComUtilities.Release(ref sourceRows);
                ComUtilities.Release(ref pivotTable);
                ComUtilities.Release(ref pivotCache);
                ComUtilities.Release(ref pivotCaches);
                ComUtilities.Release(ref destRangeObj);
                ComUtilities.Release(ref destWorksheet);
                ComUtilities.Release(ref sourceRangeObj);
                ComUtilities.Release(ref sourceWorksheet);
            }
        }, timeoutCts.Token);
    }

    /// <summary>
    /// Creates a PivotTable from an Excel Table
    /// </summary>
    public PivotTableCreateResult CreateFromTable(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName, string? shareCacheWith = null, string? shareCacheWithSheet = null)
    {
        ValidateCacheSharingInputs(shareCacheWith, shareCacheWithSheet);
        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.Worksheet? destWorksheet = null;
            Excel.Range? destRangeObj = null;
            Excel.PivotCaches? pivotCaches = null;
            Excel.PivotCache? pivotCache = null;
            Excel.PivotTable? pivotTable = null;
            Excel.Range? tableRange = null;
            Excel.Range? tableRows = null;
            Excel.Worksheet? tableSheet = null;

            try
            {
                Excel.Sheets? sheets = null;
                bool tableFound = false;

                try
                {
                    sheets = ctx.Book.Worksheets;
                    for (int i = 1; i <= sheets.Count; i++)
                    {
                        Excel.Worksheet? sheet = null;
                        Excel.ListObjects? listObjects = null;
                        try
                        {
                            sheet = (Excel.Worksheet)sheets[i];
                            listObjects = sheet.ListObjects;

                            for (int j = 1; j <= listObjects.Count; j++)
                            {
                                Excel.ListObject? tbl = null;
                                try
                                {
                                    tbl = listObjects.Item[j];
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

                if (!tableFound || table is null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found in workbook");
                }

                tableRange = table.Range;
                tableRows = tableRange.Rows;
                int rowCount = tableRows.Count;
                var headers = new List<string>();

                if (rowCount < 2)
                {
                    throw new InvalidOperationException($"Table '{tableName}' must contain at least one data row (has {rowCount} rows including header)");
                }

                // Get headers
                Excel.Range? headerRowCol = null;
                try
                {
                    headerRowCol = table.HeaderRowRange;
                    var headerValues = ExcelValueNormalizer.Normalize((object?)headerRowCol.Value2);

                    foreach (var value in headerValues.Values[0])
                    {
                        var header = value?.ToString();
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

                // Resolve the destination before Excel creates anything.
                destWorksheet = Utilities.WorksheetLookup.Get(ctx.Book, destinationSheet);
                EnsurePivotTableNameAvailable(destWorksheet, pivotTableName, destinationSheet);
                destRangeObj = destWorksheet.Range[destinationCell];

                // Create PivotCache from table
                tableSheet = (Excel.Worksheet)table.Parent;
                string sourceDataRef = $"{tableSheet.Name}!{table.Name}";

                if (shareCacheWith is not null)
                    pivotCache = ResolveSharedCreationCache(ctx.Book, ctx.App, shareCacheWithSheet!, shareCacheWith,
                        (Excel.Worksheet)tableSheet, (Excel.Range)tableRange, (Excel.ListObject)table);
                else
                {
                    pivotCaches = ctx.Book.PivotCaches();
                    pivotCache = (Excel.PivotCache)pivotCaches.Create(
                        SourceType: Excel.XlPivotTableSourceType.xlDatabase,
                        SourceData: sourceDataRef
                    );
                }

                pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: destRangeObj,
                    TableName: pivotTableName
                );

                string rangeAddress = FinishCreatedPivotTable(pivotTable, pivotTableName, destinationSheet);
                return new PivotTableCreateResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    SheetName = destinationSheet,
                    Range = rangeAddress,
                    SourceData = sourceDataRef,
                    SourceRowCount = rowCount - 1,
                    AvailableFields = headers,
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
                ComUtilities.Release(ref tableSheet);
                ComUtilities.Release(ref tableRows);
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Creates a PivotTable from a Power Pivot Data Model table
    /// Uses xlExternal source type with "ThisWorkbookDataModel" connection
    /// </summary>
    public PivotTableCreateResult CreateFromDataModel(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName)
    {
        return ExcelBatchWait.Execute(batch, TimeSpan.FromMinutes(5), (ctx, ct) =>
        {
            Excel.Model? model = null;
            Excel.ModelTable? modelTable = null;
            Excel.Worksheet? destWorksheet = null;
            Excel.Range? destRangeObj = null;
            Excel.PivotCaches? pivotCaches = null;
            Excel.PivotCache? pivotCache = null;
            Excel.PivotTable? pivotTable = null;

            try
            {
                // STEP 1: Verify Data Model exists and find the table
                // NOTE: Every workbook has a Model object, but it may be empty
                model = ctx.Book.Model;

                // Find the table in the Data Model
                Excel.ModelTables? modelTables = null;
                bool tableFound = false;
                try
                {
                    modelTables = model.ModelTables;

                    // Check if Data Model has any tables
                    if (modelTables is null || modelTables.Count == 0)
                    {
                        throw new InvalidOperationException("Data Model does not contain any tables");
                    }

                    for (int i = 1; i <= modelTables!.Count; i++)
                    {
                        Excel.ModelTable? tbl = null;
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

                if (!tableFound || modelTable is null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found in Data Model");
                }

                // Get columns from the Data Model table
                var headers = new List<string>();
                int recordCount = 0;

                try
                {
                    recordCount = modelTable.RecordCount;

                    // Get columns
                    Excel.ModelTableColumns? modelColumns = null;
                    try
                    {
                        modelColumns = modelTable.ModelTableColumns;
                        for (int i = 1; i <= modelColumns.Count; i++)
                        {
                            Excel.ModelTableColumn? column = null;
                            try
                            {
                                column = modelColumns.Item(i);
                                var colName = column.Name;
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

                // STEP 2: Resolve the destination before Excel creates anything.
                destWorksheet = Utilities.WorksheetLookup.Get(ctx.Book, destinationSheet);
                EnsurePivotTableNameAvailable(destWorksheet, pivotTableName, destinationSheet);
                destRangeObj = destWorksheet.Range[destinationCell];

                // STEP 3: Create PivotCache from Data Model
                // Using xlExternal (2) with "ThisWorkbookDataModel" connection
                pivotCaches = ctx.Book.PivotCaches();

                // xlExternal = 2
                pivotCache = pivotCaches.Create(
                    SourceType: Excel.XlPivotTableSourceType.xlExternal,
                    SourceData: "ThisWorkbookDataModel"
                );

                // STEP 4: Create PivotTable from cache, then refresh to materialize its structure
                pivotTable = pivotCache.CreatePivotTable(
                    TableDestination: destRangeObj,
                    TableName: pivotTableName
                );

                string rangeAddress = FinishCreatedPivotTable(pivotTable, pivotTableName, destinationSheet);
                return new PivotTableCreateResult
                {
                    Success = true,
                    PivotTableName = pivotTableName,
                    SheetName = destinationSheet,
                    Range = rangeAddress,
                    SourceData = $"ThisWorkbookDataModel[{tableName}]",
                    SourceRowCount = recordCount,
                    AvailableFields = headers,
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

    private static void EnsurePivotTableNameAvailable(
        Excel.Worksheet worksheet,
        string pivotTableName,
        string destinationSheet)
    {
        Excel.PivotTables? pivotTables = null;
        Excel.PivotTable? existing = null;
        try
        {
            pivotTables = (Excel.PivotTables)worksheet.PivotTables();
            int pivotTableCount = Convert.ToInt32(pivotTables.Count);
            for (int index = 1; index <= pivotTableCount; index++)
            {
                existing = pivotTables.Item(index);
                string existingName = existing.Name?.ToString() ?? string.Empty;
                if (existingName.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"PivotTable '{pivotTableName}' already exists on sheet '{destinationSheet}'. " +
                        "Choose a different PivotTable name.");
                }

                ComUtilities.Release(ref existing);
            }
        }
        finally
        {
            ComUtilities.Release(ref existing);
            ComUtilities.Release(ref pivotTables);
        }
    }

    /// <summary>
    /// Refreshes a newly created PivotTable and reads its location. If either step fails, the error
    /// names the PivotTable that remains in the workbook.
    /// </summary>
    private static string FinishCreatedPivotTable(Excel.PivotTable pivotTable, string pivotTableName, string sheetName)
    {
        Excel.Range? tableRange2 = null;
        string step = $"Refreshing the new PivotTable '{pivotTableName}'";
        try
        {
            if (!pivotTable.RefreshTable())
            {
                throw new InvalidOperationException("Excel did not refresh the new PivotTable.");
            }

            step = $"Reading the location of the new PivotTable '{pivotTableName}'";
            tableRange2 = pivotTable.TableRange2;
            return tableRange2.Address;
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
        {
            throw CreatedObjectFailure.Create("PivotTable", pivotTableName, sheetName, step, ex);
        }
        finally
        {
            ComUtilities.Release(ref tableRange2);
        }
    }
}
