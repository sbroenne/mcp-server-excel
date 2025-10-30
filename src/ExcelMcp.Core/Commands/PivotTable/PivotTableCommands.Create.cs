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
    /// </summary>
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotTableCreateResult> CreateFromRangeAsync(IExcelBatch batch,
        string sourceSheet, string sourceRange,
        string destinationSheet, string destinationCell,
        string pivotTableName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
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
                // STEP 1: Validate source data
                sourceWorksheet = ctx.Book.Worksheets.Item(sourceSheet);
                sourceRangeObj = sourceWorksheet.Range[sourceRange];

                if (sourceRangeObj.Rows.Count < 2)
                {
                    throw new InvalidOperationException($"Source range must contain headers and at least one data row. Found {sourceRangeObj.Rows.Count} rows");
                }

                // Check for headers in first row
                dynamic? headerRowRange = null;
                try
                {
                    headerRowRange = sourceRangeObj.Rows[1];
                    object[,] headerValues = headerRowRange.Value2;
                    var headers = new List<string>();
                    
                    for (int col = 1; col <= headerValues.GetLength(1); col++)
                    {
                        var header = headerValues[1, col]?.ToString();
                        if (string.IsNullOrWhiteSpace(header))
                        {
                            throw new InvalidOperationException($"Missing header in column {col}. All columns must have headers.");
                        }
                        headers.Add(header);
                    }

                    // STEP 2: Create PivotCache
                    pivotCaches = ctx.Book.PivotCaches();
                    string sourceDataRef = $"{sourceSheet}!{sourceRange}";
                    
                    // xlDatabase = 1
                    pivotCache = pivotCaches.Create(
                        SourceType: 1,
                        SourceData: sourceDataRef
                    );

                    // STEP 3: Create PivotTable
                    destWorksheet = ctx.Book.Worksheets.Item(destinationSheet);
                    destRangeObj = destWorksheet.Range[destinationCell];

                    pivotTable = pivotCache.CreatePivotTable(
                        TableDestination: destRangeObj,
                        TableName: pivotTableName
                    );

                    // STEP 4: CRITICAL - Refresh to materialize layout
                    pivotTable.RefreshTable();

                    // STEP 5: Detect field types for LLM guidance
                    var numericFields = new List<string>();
                    var textFields = new List<string>();
                    var dateFields = new List<string>();

                    // Sample first data row to detect types
                    if (sourceRangeObj.Rows.Count > 1)
                    {
                        dynamic? dataRowRange = null;
                        try
                        {
                            dataRowRange = sourceRangeObj.Rows[2];
                            object[,] dataValues = dataRowRange.Value2;

                            for (int col = 1; col <= dataValues.GetLength(1); col++)
                            {
                                string fieldName = headers[col - 1];
                                var value = dataValues[1, col];

                                if (value != null)
                                {
                                    if (DateTime.TryParse(value.ToString(), out _))
                                    {
                                        dateFields.Add(fieldName);
                                    }
                                    else if (double.TryParse(value.ToString(), out _))
                                    {
                                        numericFields.Add(fieldName);
                                    }
                                    else
                                    {
                                        textFields.Add(fieldName);
                                    }
                                }
                                else
                                {
                                    textFields.Add(fieldName);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref dataRowRange);
                        }
                    }

                    return new PivotTableCreateResult
                    {
                        Success = true,
                        PivotTableName = pivotTableName,
                        SheetName = destinationSheet,
                        Range = pivotTable.TableRange2.Address,
                        SourceData = sourceDataRef,
                        SourceRowCount = sourceRangeObj.Rows.Count - 1, // Exclude headers
                        AvailableFields = headers,
                        NumericFields = numericFields,
                        TextFields = textFields,
                        DateFields = dateFields,
                        FilePath = batch.WorkbookPath,
                        SuggestedNextActions = new List<string>
                        {
                            "Add row field(s) using AddRowFieldAsync",
                            "Add value field(s) using AddValueFieldAsync",
                            "Add filter field(s) using AddFilterFieldAsync"
                        }
                    };
                }
                finally
                {
                    ComUtilities.Release(ref headerRowRange);
                }
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
    #pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<PivotTableCreateResult> CreateFromTableAsync(IExcelBatch batch,
        string tableName,
        string destinationSheet, string destinationCell,
        string pivotTableName)
    {
        return await batch.ExecuteAsync(async (ctx, ct) =>
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
                        FilePath = batch.WorkbookPath,
                        SuggestedNextActions = new List<string>
                        {
                            "Add row field(s) using AddRowFieldAsync",
                            "Add value field(s) using AddValueFieldAsync"
                        }
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
}
