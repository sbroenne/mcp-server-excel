using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Commands.PowerQuery;

/// <summary>
/// Power Query commands - Phase 1 API (Atomic operations with explicit intent)
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Creates new query from M code file with atomic import + load operation
    /// DEFAULT: loadTo = PowerQueryLoadMode.LoadToTable (validate by executing)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name for the new query</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <param name="loadTo">Where to load the data (default: LoadToTable)</param>
    /// <param name="worksheetName">Target worksheet name (required for LoadToTable/LoadToBoth)</param>
    /// <returns>Result with query creation and data load status</returns>
    public async Task<PowerQueryCreateResult> CreateAsync(
        IExcelBatch batch,
        string queryName,
        string mCodeFile,
        PowerQueryLoadMode loadTo = PowerQueryLoadMode.LoadToTable,
        string? worksheetName = null)
    {
        var result = new PowerQueryCreateResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName,
            LoadDestination = loadTo,
            WorksheetName = worksheetName
        };

        try
        {
            // Validate inputs
            if (string.IsNullOrWhiteSpace(queryName))
            {
                result.Success = false;
                result.ErrorMessage = "Query name cannot be empty";
                return result;
            }

            if (!File.Exists(mCodeFile))
            {
                result.Success = false;
                result.ErrorMessage = $"M code file not found: {mCodeFile}";
                return result;
            }

            if ((loadTo == PowerQueryLoadMode.LoadToTable || loadTo == PowerQueryLoadMode.LoadToBoth)
                && string.IsNullOrWhiteSpace(worksheetName))
            {
                result.Success = false;
                result.ErrorMessage = "Worksheet name required when loading to table";
                return result;
            }

            // Read M code
            var mCode = await File.ReadAllTextAsync(mCodeFile);
            if (string.IsNullOrWhiteSpace(mCode))
            {
                result.Success = false;
                result.ErrorMessage = "M code file is empty";
                return result;
            }

            return await batch.Execute((ctx, ct) =>
            {
                dynamic? queries = null;
                dynamic? query = null;
                dynamic? sheet = null;
                dynamic? queryTable = null;

                try
                {
                    queries = ctx.Book.Queries;

                    // Check if query already exists
                    if (ComInterop.ComUtilities.FindQuery(ctx.Book, queryName) != null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' already exists";
                        return result;
                    }

                    // Create query with M code
                    query = queries.Add(queryName, mCode);
                    result.QueryCreated = true;

                    // Apply load destination based on mode
                    switch (loadTo)
                    {
                        case PowerQueryLoadMode.ConnectionOnly:
                            // Connection only - no data load
                            result.DataLoaded = false;
                            result.RowsLoaded = 0;
                            break;

                        case PowerQueryLoadMode.LoadToTable:
                            // Load to worksheet table
                            sheet = ctx.Book.Worksheets.Item(worksheetName!);
                            queryTable = CreateQueryTableForQuery(sheet, query);
                            queryTable.Refresh(false);  // Synchronous refresh
                            result.DataLoaded = true;
                            result.RowsLoaded = queryTable.ResultRange.Rows.Count - 1;  // Exclude header
                            break;

                        case PowerQueryLoadMode.LoadToDataModel:
                            // Load to Data Model
                            var connection = FindConnectionForQuery(ctx.Book, queryName);
                            if (connection != null)
                            {
                                connection.OLEDBConnection.Connection = $"Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                                ctx.Book.Model.DataModelConnection.RefreshAsync();
                                result.DataLoaded = true;
                                result.RowsLoaded = -1;  // Data Model doesn't expose row count easily
                            }
                            break;

                        case PowerQueryLoadMode.LoadToBoth:
                            // Load to both worksheet and Data Model
                            sheet = ctx.Book.Worksheets.Item(worksheetName!);
                            queryTable = CreateQueryTableForQuery(sheet, query);
                            queryTable.Refresh(false);
                            
                            var connection2 = FindConnectionForQuery(ctx.Book, queryName);
                            if (connection2 != null)
                            {
                                ctx.Book.Model.DataModelConnection.RefreshAsync();
                            }
                            
                            result.DataLoaded = true;
                            result.RowsLoaded = queryTable.ResultRange.Rows.Count - 1;
                            break;
                    }

                    result.Success = true;
                    result.SuggestedNextActions = loadTo switch
                    {
                        PowerQueryLoadMode.ConnectionOnly => new List<string>
                        {
                            "Use LoadToAsync() to load data when ready",
                            "Update M code with UpdateMCodeAsync() if needed",
                            "Use RefreshAsync() after applying load destination"
                        },
                        PowerQueryLoadMode.LoadToTable => new List<string>
                        {
                            $"Verify data in worksheet '{worksheetName}'",
                            "Use RefreshAsync() to reload data from source",
                            "Update M code with UpdateMCodeAsync() if needed"
                        },
                        _ => new List<string>
                        {
                            "Verify data was loaded correctly",
                            "Use RefreshAsync() to reload data from source"
                        }
                    };

                    return result;
                }
                catch (COMException ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Excel COM error creating query: {ex.Message}";
                    result.IsRetryable = ex.HResult == -2147417851;  // RPC_E_SERVERCALL_RETRYLATER
                    return result;
                }
                finally
                {
                    ComInterop.ComUtilities.Release(ref queryTable!);
                    ComInterop.ComUtilities.Release(ref sheet!);
                    ComInterop.ComUtilities.Release(ref query!);
                    ComInterop.ComUtilities.Release(ref queries!);
                }
            }, cancellationToken: default);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error creating query: {ex.Message}";
            result.IsRetryable = false;
            return result;
        }
    }

    /// <summary>
    /// Updates ONLY the M code formula (no refresh)
    /// Use RefreshAsync() separately if data update needed
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCodeFile">Path to new M code file</param>
    /// <returns>Operation result</returns>
    public async Task<OperationResult> UpdateMCodeAsync(
        IExcelBatch batch,
        string queryName,
        string mCodeFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "update-mcode"
        };

        try
        {
            if (!File.Exists(mCodeFile))
            {
                result.Success = false;
                result.ErrorMessage = $"M code file not found: {mCodeFile}";
                return result;
            }

            var mCode = await File.ReadAllTextAsync(mCodeFile);
            if (string.IsNullOrWhiteSpace(mCode))
            {
                result.Success = false;
                result.ErrorMessage = "M code file is empty";
                return result;
            }

            return await batch.Execute((ctx, ct) =>
            {
                dynamic? queries = null;
                dynamic? query = null;

                try
                {
                    queries = ctx.Book.Queries;
                    query = ComInterop.ComUtilities.FindQuery(ctx.Book, queryName);

                    if (query == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        return result;
                    }

                    // Update M code formula
                    query.Formula = mCode;
                    result.Success = true;
                    result.SuggestedNextActions = new List<string>
                    {
                        "Use RefreshAsync() to reload data with new M code",
                        "Use GetLoadConfigAsync() to check current load configuration",
                        "Verify M code syntax is valid"
                    };

                    return result;
                }
                catch (COMException ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Excel COM error updating M code: {ex.Message}";
                    result.IsRetryable = ex.HResult == -2147417851;
                    return result;
                }
                finally
                {
                    ComInterop.ComUtilities.Release(ref query!);
                    ComInterop.ComUtilities.Release(ref queries!);
                }
            }, cancellationToken: default);
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error updating M code: {ex.Message}";
            result.IsRetryable = false;
            return result;
        }
    }

    /// <summary>
    /// Helper method to create QueryTable for a query
    /// </summary>
    private dynamic CreateQueryTableForQuery(dynamic sheet, dynamic query)
    {
        string queryName = query.Name;
        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
        
        dynamic range = sheet.Range["A1"];
        dynamic queryTable = sheet.QueryTables.Add(connectionString, range, $"Table_{queryName}");
        
        queryTable.Name = $"Query_{queryName}";
        queryTable.RefreshStyle = 1;  // xlInsertDeleteCells
        queryTable.RowNumbers = false;
        queryTable.FillAdjacentFormulas = false;
        queryTable.PreserveFormatting = true;
        queryTable.RefreshOnFileOpen = false;
        queryTable.BackgroundQuery = false;  // Synchronous refresh
        queryTable.SavePassword = false;
        queryTable.SaveData = true;
        queryTable.AdjustColumnWidth = true;
        queryTable.RefreshPeriod = 0;
        queryTable.PreserveColumnInfo = true;

        return queryTable;
    }

    /// <summary>
    /// Helper method to find connection for a query
    /// </summary>
    private dynamic? FindConnectionForQuery(dynamic workbook, string queryName)
    {
        dynamic? connections = null;
        try
        {
            connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = null;
                try
                {
                    conn = connections.Item(i);
                    string connName = conn.Name;
                    if (connName.Contains(queryName))
                    {
                        return conn;
                    }
                }
                finally
                {
                    if (conn != null && conn != connections.Item(i))
                    {
                        ComInterop.ComUtilities.Release(ref conn!);
                    }
                }
            }
        }
        finally
        {
            ComInterop.ComUtilities.Release(ref connections!);
        }
        
        return null;
    }
}
