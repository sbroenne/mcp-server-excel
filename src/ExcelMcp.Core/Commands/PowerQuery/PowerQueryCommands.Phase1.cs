using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Commands;

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

            // Default to query name for worksheet name (Excel's default behavior)
            if ((loadTo == PowerQueryLoadMode.LoadToTable || loadTo == PowerQueryLoadMode.LoadToBoth)
                && string.IsNullOrWhiteSpace(worksheetName))
            {
                worksheetName = queryName;
                result.WorksheetName = worksheetName;
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
                            // Load to worksheet table - create sheet if it doesn't exist
                            dynamic? worksheets = null;
                            try
                            {
                                worksheets = ctx.Book.Worksheets;
                                try
                                {
                                    sheet = worksheets.Item(worksheetName!);
                                }
                                catch (System.Runtime.InteropServices.COMException)
                                {
                                    // Sheet doesn't exist, create it
                                    sheet = worksheets.Add();
                                    sheet.Name = worksheetName;
                                }
                            }
                            finally
                            {
                                ComInterop.ComUtilities.Release(ref worksheets!);
                            }
                            
                            queryTable = CreateQueryTableForQuery(sheet, query);
                            queryTable.Refresh(false);  // Synchronous refresh
                            result.DataLoaded = true;
                            result.RowsLoaded = queryTable.ResultRange.Rows.Count - 1;  // Exclude header
                            break;

                        case PowerQueryLoadMode.LoadToDataModel:
                            // Load to Data Model using Connections.Add2 method
                            dynamic? connections = null;
                            dynamic? dmConnection = null;
                            try
                            {
                                connections = ctx.Book.Connections;
                                string connectionName = $"Query - {queryName}";
                                string description = $"Connection to the '{queryName}' query in the workbook.";
                                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                                string commandText = $"\"{queryName}\"";
                                int commandType = 6; // Data Model command type
                                bool createModelConnection = true; // CRITICAL: This loads to Data Model
                                bool importRelationships = false;

                                dmConnection = connections.Add2(
                                    connectionName,
                                    description,
                                    connectionString,
                                    commandText,
                                    commandType,
                                    createModelConnection,
                                    importRelationships
                                );
                                result.DataLoaded = true;
                                result.RowsLoaded = -1;  // Data Model doesn't expose row count easily
                            }
                            finally
                            {
                                ComInterop.ComUtilities.Release(ref dmConnection!);
                                ComInterop.ComUtilities.Release(ref connections!);
                            }
                            break;

                        case PowerQueryLoadMode.LoadToBoth:
                            // Load to both worksheet and Data Model - create sheet if it doesn't exist
                            dynamic? worksheetsBoth = null;
                            try
                            {
                                worksheetsBoth = ctx.Book.Worksheets;
                                try
                                {
                                    sheet = worksheetsBoth.Item(worksheetName!);
                                }
                                catch (System.Runtime.InteropServices.COMException)
                                {
                                    // Sheet doesn't exist, create it
                                    sheet = worksheetsBoth.Add();
                                    sheet.Name = worksheetName;
                                }
                            }
                            finally
                            {
                                ComInterop.ComUtilities.Release(ref worksheetsBoth!);
                            }
                            
                            queryTable = CreateQueryTableForQuery(sheet, query);
                            queryTable.Refresh(false);

                            // Also load to Data Model
                            dynamic? connectionsBoth = null;
                            dynamic? dmConnectionBoth = null;
                            try
                            {
                                connectionsBoth = ctx.Book.Connections;
                                string connectionName = $"Query - {queryName}";
                                string description = $"Connection to the '{queryName}' query in the workbook.";
                                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                                string commandText = $"\"{queryName}\"";
                                int commandType = 6;
                                bool createModelConnection = true;
                                bool importRelationships = false;

                                dmConnectionBoth = connectionsBoth.Add2(
                                    connectionName,
                                    description,
                                    connectionString,
                                    commandText,
                                    commandType,
                                    createModelConnection,
                                    importRelationships
                                );
                            }
                            finally
                            {
                                ComInterop.ComUtilities.Release(ref dmConnectionBoth!);
                                ComInterop.ComUtilities.Release(ref connectionsBoth!);
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
    /// Sets query load destination and refreshes data atomically
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="loadTo">Where to load the data</param>
    /// <param name="worksheetName">Target worksheet (required for LoadToTable/LoadToBoth)</param>
    /// <returns>Result with load configuration and refresh status</returns>
    public async Task<PowerQueryLoadResult> LoadToAsync(
        IExcelBatch batch,
        string queryName,
        PowerQueryLoadMode loadTo,
        string? worksheetName = null)
    {
        var result = new PowerQueryLoadResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName,
            LoadDestination = loadTo,
            WorksheetName = worksheetName
        };

        try
        {
            if ((loadTo == PowerQueryLoadMode.LoadToTable || loadTo == PowerQueryLoadMode.LoadToBoth)
                && string.IsNullOrWhiteSpace(worksheetName))
            {
                result.Success = false;
                result.ErrorMessage = "Worksheet name required for LoadToTable/LoadToBoth";
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
                    query = ComInterop.ComUtilities.FindQuery(ctx.Book, queryName);

                    if (query == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        return result;
                    }

                    // Apply load destination
                    switch (loadTo)
                    {
                        case PowerQueryLoadMode.LoadToTable:
                            sheet = ctx.Book.Worksheets.Item(worksheetName!);
                            queryTable = CreateQueryTableForQuery(sheet, query);
                            queryTable.Refresh(false);
                            result.ConfigurationApplied = true;
                            result.DataRefreshed = true;
                            result.RowsLoaded = queryTable.ResultRange.Rows.Count - 1;
                            break;

                        case PowerQueryLoadMode.LoadToDataModel:
                            // Load to Data Model using Connections.Add2 method
                            dynamic? connectionsLoadTo = null;
                            dynamic? dmConnectionLoadTo = null;
                            try
                            {
                                connectionsLoadTo = ctx.Book.Connections;
                                string connectionName = $"Query - {queryName}";
                                string description = $"Connection to the '{queryName}' query in the workbook.";
                                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                                string commandText = $"\"{queryName}\"";
                                int commandType = 6;
                                bool createModelConnection = true;
                                bool importRelationships = false;

                                dmConnectionLoadTo = connectionsLoadTo.Add2(
                                    connectionName,
                                    description,
                                    connectionString,
                                    commandText,
                                    commandType,
                                    createModelConnection,
                                    importRelationships
                                );
                                result.ConfigurationApplied = true;
                                result.DataRefreshed = true;
                                result.RowsLoaded = -1;
                            }
                            finally
                            {
                                ComInterop.ComUtilities.Release(ref dmConnectionLoadTo!);
                                ComInterop.ComUtilities.Release(ref connectionsLoadTo!);
                            }
                            break;

                        case PowerQueryLoadMode.LoadToBoth:
                            sheet = ctx.Book.Worksheets.Item(worksheetName!);
                            queryTable = CreateQueryTableForQuery(sheet, query);
                            queryTable.Refresh(false);

                            // Also load to Data Model
                            dynamic? connectionsLoadToBoth = null;
                            dynamic? dmConnectionLoadToBoth = null;
                            try
                            {
                                connectionsLoadToBoth = ctx.Book.Connections;
                                string connectionName = $"Query - {queryName}";
                                string description = $"Connection to the '{queryName}' query in the workbook.";
                                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                                string commandText = $"\"{queryName}\"";
                                int commandType = 6;
                                bool createModelConnection = true;
                                bool importRelationships = false;

                                dmConnectionLoadToBoth = connectionsLoadToBoth.Add2(
                                    connectionName,
                                    description,
                                    connectionString,
                                    commandText,
                                    commandType,
                                    createModelConnection,
                                    importRelationships
                                );
                            }
                            finally
                            {
                                ComInterop.ComUtilities.Release(ref dmConnectionLoadToBoth!);
                                ComInterop.ComUtilities.Release(ref connectionsLoadToBoth!);
                            }

                            result.ConfigurationApplied = true;
                            result.DataRefreshed = true;
                            result.RowsLoaded = queryTable.ResultRange.Rows.Count - 1;
                            break;

                        case PowerQueryLoadMode.ConnectionOnly:
                            result.ConfigurationApplied = true;
                            result.DataRefreshed = false;
                            result.RowsLoaded = 0;
                            break;
                    }

                    result.Success = true;
                    result.SuggestedNextActions = new List<string>
                    {
                        "Verify data was loaded correctly",
                        "Use RefreshAsync() to reload data from source",
                        "Update M code with UpdateMCodeAsync() if needed"
                    };

                    return result;
                }
                catch (COMException ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Excel COM error applying load destination: {ex.Message}";
                    result.IsRetryable = ex.HResult == -2147417851;
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
            result.ErrorMessage = $"Error applying load destination: {ex.Message}";
            result.IsRetryable = false;
            return result;
        }
    }

    /// <summary>
    /// Converts query to connection-only (removes data load)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <returns>Operation result</returns>
    public async Task<OperationResult> UnloadAsync(
        IExcelBatch batch,
        string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "unload"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;

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

                // Remove QueryTables from all worksheets
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    dynamic? queryTables = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        queryTables = sheet.QueryTables;

                        for (int j = queryTables.Count; j >= 1; j--)
                        {
                            dynamic? qt = null;
                            try
                            {
                                qt = queryTables.Item(j);
                                string qtName = qt.Name;
                                if (qtName.Contains(queryName))
                                {
                                    qt.Delete();
                                }
                            }
                            finally
                            {
                                ComInterop.ComUtilities.Release(ref qt!);
                            }
                        }
                    }
                    finally
                    {
                        ComInterop.ComUtilities.Release(ref queryTables!);
                        ComInterop.ComUtilities.Release(ref sheet!);
                    }
                }

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    "Query is now connection-only",
                    "Use LoadToAsync() to load data when ready",
                    "M code is preserved and can be edited"
                };

                return result;
            }
            catch (COMException ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Excel COM error removing data load: {ex.Message}";
                result.IsRetryable = ex.HResult == -2147417851;
                return result;
            }
            finally
            {
                ComInterop.ComUtilities.Release(ref sheets!);
                ComInterop.ComUtilities.Release(ref query!);
                ComInterop.ComUtilities.Release(ref queries!);
            }
        }, cancellationToken: default);
    }

    // ValidateSyntaxAsync removed - Excel doesn't validate M code syntax at query creation time.
    // Validation only happens during refresh, making syntax-only validation unreliable.
    // Users should use CreateAsync + RefreshAsync to discover syntax errors.

    /// <summary>
    /// Updates M code and refreshes data in one atomic operation (convenience method)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <param name="mCodeFile">Path to new M code file</param>
    /// <returns>Operation result</returns>
    public async Task<OperationResult> UpdateAndRefreshAsync(
        IExcelBatch batch,
        string queryName,
        string mCodeFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "update-and-refresh"
        };

        try
        {
            // Update M code
            var updateResult = await UpdateMCodeAsync(batch, queryName, mCodeFile);
            if (!updateResult.Success)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to update M code: {updateResult.ErrorMessage}";
                return result;
            }

            // Refresh data
            var refreshResult = await RefreshAsync(batch, queryName);
            if (!refreshResult.Success)
            {
                result.Success = false;
                result.ErrorMessage = $"M code updated but refresh failed: {refreshResult.ErrorMessage}";
                return result;
            }

            result.Success = true;
            result.SuggestedNextActions = new List<string>
            {
                "M code updated and data refreshed successfully",
                "Verify data matches expectations",
                "Use GetLoadConfigAsync() to check load configuration"
            };

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error updating and refreshing: {ex.Message}";
            result.IsRetryable = false;
            return result;
        }
    }

    /// <summary>
    /// Refreshes all Power Query queries in the workbook
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>Operation result with refresh summary</returns>
    public async Task<OperationResult> RefreshAllAsync(IExcelBatch batch)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;

            try
            {
                queries = ctx.Book.Queries;
                int totalQueries = queries.Count;
                int refreshedCount = 0;
                var errors = new List<string>();

                for (int i = 1; i <= totalQueries; i++)
                {
                    dynamic? query = null;
                    try
                    {
                        query = queries.Item(i);
                        string queryName = query.Name;

                        // Refresh via connection
                        var connection = FindConnectionForQuery(ctx.Book, queryName);
                        if (connection != null)
                        {
                            try
                            {
                                connection.Refresh();
                                refreshedCount++;
                            }
                            catch (COMException ex)
                            {
                                errors.Add($"{queryName}: {ex.Message}");
                            }
                        }
                    }
                    finally
                    {
                        ComInterop.ComUtilities.Release(ref query!);
                    }
                }

                // ✅ Rule 0: Success = false when errors exist
                if (errors.Any())
                {
                    result.Success = false;
                    result.ErrorMessage = $"Some queries failed to refresh: {string.Join(", ", errors)}";
                }
                else
                {
                    result.Success = true;
                }

                result.SuggestedNextActions = new List<string>
                {
                    $"Refreshed {refreshedCount} of {totalQueries} queries successfully",
                    errors.Any() ? $"{errors.Count} queries had errors" : "All queries refreshed",
                    "Use RefreshAsync() to refresh individual queries"
                };

                return result;
            }
            catch (COMException ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Excel COM error refreshing queries: {ex.Message}";
                result.IsRetryable = ex.HResult == -2147417851;
                return result;
            }
            finally
            {
                ComInterop.ComUtilities.Release(ref queries!);
            }
        }, cancellationToken: default);
    }

    /// <summary>
    /// Helper method to create QueryTable for a query
    /// </summary>
    private dynamic CreateQueryTableForQuery(dynamic sheet, dynamic query)
    {
        string queryName = query.Name;
        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";

        dynamic range = sheet.Range["A1"];
        // Use Type.Missing for 3rd parameter (working pattern from diagnostic tests)
        dynamic queryTable = sheet.QueryTables.Add(connectionString, range, Type.Missing);

        queryTable.Name = queryName;
        queryTable.CommandText = $"SELECT * FROM [{queryName}]";  // Set AFTER creation (working pattern)
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

        // Note: Caller is responsible for calling Refresh(false) after QueryTable is returned
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
