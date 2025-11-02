using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query advanced operations (LoadTo, Sources, Test, Peek, Eval)
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> LoadToAsync(IExcelBatch batch, string queryName, string sheetName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-loadto"
        };

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Find or create target sheet
                dynamic? sheets = null;
                dynamic? targetSheet = null;
                try
                {
                    sheets = ctx.Book.Worksheets;

                    for (int i = 1; i <= sheets.Count; i++)
                    {
                        dynamic? sheet = null;
                        try
                        {
                            sheet = sheets.Item(i);
                            if (sheet.Name == sheetName)
                            {
                                targetSheet = sheet;
                                sheet = null; // Don't release - we're using it
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref sheet);
                        }
                    }

                    if (targetSheet == null)
                    {
                        targetSheet = sheets.Add();
                        targetSheet.Name = sheetName;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref sheets);
                }

                // Get the workbook connections to find our query
                dynamic? connections = null;
                dynamic? targetConnection = null;
                try
                {
                    connections = ctx.Book.Connections;

                    // Look for existing connection for this query
                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(i);
                            string connName = conn.Name?.ToString() ?? "";
                            if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                                connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                targetConnection = conn;
                                conn = null; // Don't release - we're using it
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                // If no connection exists, we need to create one by loading the query to table
                if (targetConnection == null)
                {
                    // Access the query through the Queries collection and load it to table
                    dynamic? queries = null;
                    dynamic? targetQuery = null;
                    dynamic? queryTables = null;
                    dynamic? queryTable = null;
                    dynamic? rangeObj = null;
                    try
                    {
                        queries = ctx.Book.Queries;

                        for (int i = 1; i <= queries.Count; i++)
                        {
                            dynamic? q = null;
                            try
                            {
                                q = queries.Item(i);
                                if (q.Name.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                                {
                                    targetQuery = q;
                                    q = null; // Don't release - we're using it
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref q);
                            }
                        }

                        if (targetQuery == null)
                        {
                            result.Success = false;
                            result.ErrorMessage = $"Query '{queryName}' not found in queries collection";
                            return result;
                        }

                        // Create a QueryTable using the Mashup provider
                        queryTables = targetSheet.QueryTables;
                        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                        string commandText = $"SELECT * FROM [{queryName}]";

                        rangeObj = targetSheet.Range["A1"];
                        queryTable = queryTables.Add(connectionString, rangeObj, commandText);
                        queryTable.Name = queryName.Replace(" ", "_");
                        queryTable.RefreshStyle = 1; // xlInsertDeleteCells

                        // Set additional properties for better data loading
                        queryTable.BackgroundQuery = false; // Don't run in background
                        queryTable.PreserveColumnInfo = true;
                        queryTable.PreserveFormatting = true;
                        queryTable.AdjustColumnWidth = true;

                        // Refresh to actually load the data
                        queryTable.Refresh(false); // false = wait for completion
                    }
                    finally
                    {
                        ComUtilities.Release(ref rangeObj);
                        ComUtilities.Release(ref queryTable);
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref targetQuery);
                        ComUtilities.Release(ref queries);
                    }
                }
                else
                {
                    // Connection exists, create QueryTable from existing connection
                    dynamic? queryTables = null;
                    dynamic? queryTable = null;
                    dynamic? rangeObj = null;
                    try
                    {
                        queryTables = targetSheet.QueryTables;

                        // Remove any existing QueryTable with the same name
                        try
                        {
                            for (int i = queryTables.Count; i >= 1; i--)
                            {
                                dynamic? qt = null;
                                try
                                {
                                    qt = queryTables.Item(i);
                                    if (qt.Name.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase))
                                    {
                                        qt.Delete();
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref qt);
                                }
                            }
                        }
                        catch { } // Ignore errors if no existing QueryTable

                        // Create new QueryTable
                        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                        string commandText = $"SELECT * FROM [{queryName}]";

                        rangeObj = targetSheet.Range["A1"];
                        queryTable = queryTables.Add(connectionString, rangeObj, commandText);
                        queryTable.Name = queryName.Replace(" ", "_");
                        queryTable.RefreshStyle = 1; // xlInsertDeleteCells
                        queryTable.BackgroundQuery = false;
                        queryTable.PreserveColumnInfo = true;
                        queryTable.PreserveFormatting = true;
                        queryTable.AdjustColumnWidth = true;

                        // Refresh to load data
                        queryTable.Refresh(false);
                    }
                    finally
                    {
                        ComUtilities.Release(ref rangeObj);
                        ComUtilities.Release(ref queryTable);
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref targetConnection);
                    }
                }

                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref query);
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error loading query to worksheet: {ex.Message}";
                return result;
            }
        });
    }

    /// <inheritdoc />
    public async Task<WorksheetListResult> SourcesAsync(IExcelBatch batch)
    {
        var result = new WorksheetListResult { FilePath = batch.WorkbookPath };

        return await batch.Execute<WorksheetListResult>((ctx, ct) =>
        {
            dynamic? worksheets = null;
            dynamic? names = null;
            try
            {
                // Get all tables from all worksheets
                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    dynamic? tables = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);
                        string wsName = worksheet.Name;

                        tables = worksheet.ListObjects;
                        for (int i = 1; i <= tables.Count; i++)
                        {
                            dynamic? table = null;
                            try
                            {
                                table = tables.Item(i);
                                result.Worksheets.Add(new WorksheetInfo
                                {
                                    Name = table.Name,
                                    Index = i,
                                    Visible = true
                                });
                            }
                            finally
                            {
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                // Get all named ranges
                names = ctx.Book.Names;
                int namedRangeIndex = result.Worksheets.Count + 1;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic? name = null;
                    try
                    {
                        name = names.Item(i);
                        string nameValue = name.Name;
                        if (!nameValue.StartsWith("_"))
                        {
                            result.Worksheets.Add(new WorksheetInfo
                            {
                                Name = nameValue,
                                Index = namedRangeIndex++,
                                Visible = true
                            });
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref name);
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing sources: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref worksheets);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> TestAsync(IExcelBatch batch, string sourceName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-test"
        };

        return await batch.Execute<OperationResult>((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            dynamic? tempQuery = null;
            try
            {
                // Create a test query to load the source
                string testQuery = $@"
let
    Source = Excel.CurrentWorkbook(){{[Name=""{sourceName.Replace("\"", "\"\"")}""]]}}[Content]
in
    Source";

                queriesCollection = ctx.Book.Queries;
                tempQuery = queriesCollection.Add("_TestQuery", testQuery);

                // Try to refresh
                bool refreshSuccess = false;
                try
                {
                    tempQuery.Refresh();
                    refreshSuccess = true;
                }
                catch { }

                // Clean up
                tempQuery.Delete();

                // If refresh failed, this is a failure
                result.Success = refreshSuccess;
                if (!refreshSuccess)
                {
                    result.ErrorMessage = "Source exists but could not refresh (may need data source configuration)";
                }

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Source '{sourceName}' not found or cannot be loaded: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref tempQuery);
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

    /// <inheritdoc />
    public async Task<WorksheetDataResult> PeekAsync(IExcelBatch batch, string sourceName)
    {
        var result = new WorksheetDataResult
        {
            FilePath = batch.WorkbookPath,
            SheetName = sourceName
        };

        return await batch.Execute<WorksheetDataResult>((ctx, ct) =>
        {
            dynamic? names = null;
            dynamic? worksheets = null;
            try
            {
                // Check if it's a named range (single value)
                names = ctx.Book.Names;
                for (int i = 1; i <= names.Count; i++)
                {
                    dynamic? name = null;
                    try
                    {
                        name = names.Item(i);
                        string nameValue = name.Name;
                        if (nameValue == sourceName)
                        {
                            try
                            {
                                var value = name.RefersToRange.Value;
                                result.Data.Add([value]);
                                result.RowCount = 1;
                                result.ColumnCount = 1;
                                result.Success = true;
                                return result;
                            }
                            catch
                            {
                                result.Success = false;
                                result.ErrorMessage = "Named range found but value cannot be read (may be #REF!)";
                                return result;
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref name);
                    }
                }

                // Check if it's a table
                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    dynamic? tables = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);
                        tables = worksheet.ListObjects;
                        for (int i = 1; i <= tables.Count; i++)
                        {
                            dynamic? table = null;
                            dynamic? listCols = null;
                            try
                            {
                                table = tables.Item(i);
                                if (table.Name == sourceName)
                                {
                                    result.RowCount = table.ListRows.Count;
                                    result.ColumnCount = table.ListColumns.Count;

                                    // Get column names
                                    listCols = table.ListColumns;
                                    for (int c = 1; c <= Math.Min(result.ColumnCount, 10); c++)
                                    {
                                        dynamic? listCol = null;
                                        try
                                        {
                                            listCol = listCols.Item(c);
                                            result.Headers.Add(listCol.Name);
                                        }
                                        finally
                                        {
                                            ComUtilities.Release(ref listCol);
                                        }
                                    }

                                    result.Success = true;
                                    return result;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref listCols);
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref worksheet);
                    }
                }

                result.Success = false;
                result.ErrorMessage = $"Source '{sourceName}' not found";
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error peeking source: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref names);
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> EvalAsync(IExcelBatch batch, string mExpression)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = "_EvalExpression"
        };

        return await batch.Execute<PowerQueryViewResult>((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            dynamic? tempQuery = null;
            try
            {
                // Create a temporary query with the expression
                string evalQuery = $@"
let
    Result = {mExpression}
in
    Result";

                queriesCollection = ctx.Book.Queries;
                tempQuery = queriesCollection.Add("_EvalQuery", evalQuery);

                result.MCode = evalQuery;
                result.CharacterCount = evalQuery.Length;

                // Try to refresh
                try
                {
                    tempQuery.Refresh();
                    result.Success = true;
                    result.ErrorMessage = null;
                }
                catch (Exception refreshEx)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Expression syntax is valid but refresh failed: {refreshEx.Message}";
                }

                // Clean up
                tempQuery.Delete();

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Expression evaluation failed: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref tempQuery);
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

}
