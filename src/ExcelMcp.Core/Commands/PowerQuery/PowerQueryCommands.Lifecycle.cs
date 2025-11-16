using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query lifecycle operations (List, View, Import, Export, Update, Delete)
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public PowerQueryListResult List(IExcelBatch batch)
    {
        var result = new PowerQueryListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? queriesCollection = null;
            try
            {
                queriesCollection = ctx.Book.Queries;
                int count = queriesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic? query = null;
                    try
                    {
                        query = queriesCollection.Item(i);
                        string name = query.Name ?? $"Query{i}";
                        string formula = query.Formula ?? "";

                        string preview = formula.Length > 80 ? formula[..77] + "..." : formula;

                        // Check if connection only
                        bool isConnectionOnly = true;
                        dynamic? connections = null;
                        try
                        {
                            connections = ctx.Book.Connections;
                            for (int c = 1; c <= connections.Count; c++)
                            {
                                dynamic? conn = null;
                                try
                                {
                                    conn = connections.Item(c);
                                    string connName = conn.Name?.ToString() ?? "";
                                    if (connName.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                                        connName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase))
                                    {
                                        isConnectionOnly = false;
                                        break;
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref conn);
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            ComUtilities.Release(ref connections);
                        }

                        result.Queries.Add(new PowerQueryInfo
                        {
                            Name = name,
                            Formula = formula,
                            FormulaPreview = preview,
                            IsConnectionOnly = isConnectionOnly
                        });
                    }
                    catch (Exception queryEx)
                    {
                        result.Queries.Add(new PowerQueryInfo
                        {
                            Name = $"Error Query {i}",
                            Formula = "",
                            FormulaPreview = $"Error: {queryEx.Message}",
                            IsConnectionOnly = false
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref query);
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error accessing Power Queries: {ex.Message}";

                string extension = Path.GetExtension(batch.WorkbookPath).ToLowerInvariant();
                if (extension == ".xls")
                {
                    result.ErrorMessage += " (.xls files don't support Power Query)";
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

    /// <inheritdoc />
    public PowerQueryViewResult View(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    var queryNames = GetQueryNames(ctx.Book);
                    string? suggestion = FindClosestMatch(queryName, queryNames);

                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        result.ErrorMessage += $". Did you mean '{suggestion}'?";
                    }
                    return result;
                }

                string mCode = query.Formula;
                result.MCode = mCode;
                result.CharacterCount = mCode.Length;

                // Check if connection only
                bool isConnectionOnly = true;
                dynamic? connections = null;
                try
                {
                    connections = ctx.Book.Connections;
                    for (int c = 1; c <= connections.Count; c++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(c);
                            string connName = conn.Name?.ToString() ?? "";
                            if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                                connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                isConnectionOnly = false;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }
                }
                catch { }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                result.IsConnectionOnly = isConnectionOnly;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error viewing query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public PowerQueryLoadConfigResult GetLoadConfig(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryLoadConfigResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? worksheets = null;
            dynamic? connections = null;
            dynamic? names = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                // Check for QueryTables first (table loading)
                bool hasTableConnection = false;
                bool hasDataModelConnection = false;
                string? targetSheet = null;

                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    dynamic? queryTables = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);
                        queryTables = worksheet.QueryTables;

                        for (int qt = 1; qt <= queryTables.Count; qt++)
                        {
                            dynamic? queryTable = null;
                            try
                            {
                                queryTable = queryTables.Item(qt);
                                string qtName = queryTable.Name?.ToString() ?? "";

                                // Check if this QueryTable is for our query
                                if (qtName.Equals(queryName.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase) ||
                                    qtName.Contains(queryName.Replace(" ", "_")))
                                {
                                    hasTableConnection = true;
                                    targetSheet = worksheet.Name;
                                    ComUtilities.Release(ref queryTable);
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref queryTable);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref queryTables);
                        ComUtilities.Release(ref worksheet);
                    }
                    if (hasTableConnection) break;
                }

                // Check for connections (for data model or other types)
                connections = ctx.Book.Connections;
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
                            result.HasConnection = true;

                            // If we don't have a table connection but have a workbook connection,
                            // it's likely a data model connection
                            if (!hasTableConnection)
                            {
                                hasDataModelConnection = true;
                            }
                        }
                        else if (connName.Equals($"DataModel_{queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            // This is our explicit data model connection marker
                            result.HasConnection = true;
                            hasDataModelConnection = true;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref conn);
                    }
                }

                // Always check for named range markers that indicate data model loading
                // (even if we have table connections, for LoadToBoth mode)
                if (!hasDataModelConnection)
                {
                    // Check for our data model marker
                    try
                    {
                        names = ctx.Book.Names;
                        string markerName = $"DataModel_Query_{queryName}";

                        for (int i = 1; i <= names.Count; i++)
                        {
                            dynamic? existingName = null;
                            try
                            {
                                existingName = names.Item(i);
                                if (existingName.Name.ToString() == markerName)
                                {
                                    hasDataModelConnection = true;
                                    ComUtilities.Release(ref existingName);
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref existingName);
                            }
                        }
                    }
                    catch
                    {
                        // Cannot check names
                    }

                    // Fallback: Check if the query has data model indicators
                    if (!hasDataModelConnection)
                    {
                        hasDataModelConnection = CheckQueryDataModelConfiguration(query);
                    }
                }

                // Determine load mode
                if (hasTableConnection && hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToBoth;
                }
                else if (hasTableConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToTable;
                }
                else if (hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToDataModel;
                }
                else
                {
                    result.LoadMode = PowerQueryLoadMode.ConnectionOnly;
                }

                result.TargetSheet = targetSheet;
                result.IsLoadedToDataModel = hasDataModelConnection;
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error getting load config: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "pq-delete"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? queriesCollection = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Query '{queryName}' not found";
                    return result;
                }

                queriesCollection = ctx.Book.Queries;
                queriesCollection.Item(queryName).Delete();

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting query: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref queriesCollection);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <summary>
    /// Helper to get all query names
    /// </summary>
    private static List<string> GetQueryNames(dynamic workbook)
    {
        var names = new List<string>();
        dynamic? queriesCollection = null;
        try
        {
            queriesCollection = workbook.Queries;
            for (int i = 1; i <= queriesCollection.Count; i++)
            {
                dynamic? query = null;
                try
                {
                    query = queriesCollection.Item(i);
                    names.Add(query.Name);
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }
        }
        catch { }
        finally
        {
            ComUtilities.Release(ref queriesCollection);
        }
        return names;
    }

    /// <summary>
    /// Creates new query from M code file with atomic import + load operation
    /// DEFAULT: loadMode = PowerQueryLoadMode.LoadToTable (validate by executing)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name for the new query</param>
    /// <param name="mCodeFile">Path to M code file</param>
    /// <param name="loadMode">Where to load the data (default: LoadToTable)</param>
    /// <param name="targetSheet">Target worksheet name (required for LoadToTable/LoadToBoth)</param>
    /// <returns>Result with query creation and data load status</returns>
    public PowerQueryCreateResult Create(
        IExcelBatch batch,
        string queryName,
        string mCodeFile,
        PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable,
        string? targetSheet = null)
    {
        var result = new PowerQueryCreateResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName,
            LoadDestination = loadMode,
            WorksheetName = targetSheet
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
            if ((loadMode == PowerQueryLoadMode.LoadToTable || loadMode == PowerQueryLoadMode.LoadToBoth)
                && string.IsNullOrWhiteSpace(targetSheet))
            {
                targetSheet = queryName;
                result.WorksheetName = targetSheet;
            }

            // Read M code
            var mCode = File.ReadAllText(mCodeFile);
            if (string.IsNullOrWhiteSpace(mCode))
            {
                result.Success = false;
                result.ErrorMessage = "M code file is empty";
                return result;
            }

            return batch.Execute((ctx, ct) =>
            {
                dynamic? queries = null;
                dynamic? query = null;
                dynamic? sheet = null;
                dynamic? queryTable = null;

                try
                {
                    queries = ctx.Book.Queries;

                    // Check if query already exists
                    if (ComUtilities.FindQuery(ctx.Book, queryName) != null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' already exists";
                        return result;
                    }

                    // Create query with M code
                    query = queries.Add(queryName, mCode);
                    result.QueryCreated = true;

                    // Apply load destination based on mode
                    switch (loadMode)
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
                                    sheet = worksheets.Item(targetSheet!);
                                }
                                catch (COMException)
                                {
                                    // Sheet doesn't exist, create it
                                    sheet = worksheets.Add();
                                    sheet.Name = targetSheet;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref worksheets!);
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
                                ComUtilities.Release(ref dmConnection!);
                                ComUtilities.Release(ref connections!);
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
                                    sheet = worksheetsBoth.Item(targetSheet!);
                                }
                                catch (COMException)
                                {
                                    // Sheet doesn't exist, create it
                                    sheet = worksheetsBoth.Add();
                                    sheet.Name = targetSheet;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref worksheetsBoth!);
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
                                ComUtilities.Release(ref dmConnectionBoth!);
                                ComUtilities.Release(ref connectionsBoth!);
                            }

                            result.DataLoaded = true;
                            result.RowsLoaded = queryTable.ResultRange.Rows.Count - 1;
                            break;
                    }

                    result.Success = true;

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
                    ComUtilities.Release(ref queryTable!);
                    ComUtilities.Release(ref sheet!);
                    ComUtilities.Release(ref query!);
                    ComUtilities.Release(ref queries!);
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
    /// Updates M code and refreshes data atomically
    /// Complete operation: Updates query formula AND reloads fresh data
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query to update</param>
    /// <param name="mCodeFile">Path to new M code file</param>
    /// <returns>Operation result</returns>
    public OperationResult Update(
        IExcelBatch batch,
        string queryName,
        string mCodeFile)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "update"
        };

        try
        {
            if (!File.Exists(mCodeFile))
            {
                result.Success = false;
                result.ErrorMessage = $"M code file not found: {mCodeFile}";
                return result;
            }

            var mCode = File.ReadAllText(mCodeFile);
            if (string.IsNullOrWhiteSpace(mCode))
            {
                result.Success = false;
                result.ErrorMessage = "M code file is empty";
                return result;
            }

            return batch.Execute((ctx, ct) =>
            {
                dynamic? queries = null;
                dynamic? query = null;

                try
                {
                    queries = ctx.Book.Queries;
                    query = ComUtilities.FindQuery(ctx.Book, queryName);

                    if (query == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        return result;
                    }

                    // Update M code formula - CRITICAL: Must completely replace, not append
                    // Delete and recreate to ensure clean replacement (avoid merging bug)
                    string originalName = query.Name;

                    // Delete the old query
                    query.Delete();
                    ComUtilities.Release(ref query);
                    query = null;

                    // Create new query with updated M code
                    query = queries.Add(originalName, mCode);

                    // Auto-refresh to keep data in sync with new M code

                    // For UpdateAsync, we need to recreate QueryTables to handle column structure changes

                    // Step 1: Recreate QueryTables with new schema (handles column structure changes)
                    bool queryTableRecreated = false;
                    dynamic? sheets = null;
                    try
                    {
                        sheets = ctx.Book.Worksheets;
                        for (int s = 1; s <= sheets.Count; s++)
                        {
                            dynamic? sheet = null;
                            dynamic? queryTables = null;
                            try
                            {
                                sheet = sheets.Item(s);
                                queryTables = sheet.QueryTables;

                                // Find QueryTable for this query and recreate it
                                for (int q = queryTables.Count; q >= 1; q--)
                                {
                                    dynamic? qt = null;
                                    try
                                    {
                                        qt = queryTables.Item(q);
                                        string qtName = qt.Name?.ToString() ?? "";
                                        // Use Contains like DeleteAsync does (Excel may modify QueryTable names)
                                        if (qtName.Contains(queryName, StringComparison.OrdinalIgnoreCase))
                                        {
                                            // Delete old QueryTable
                                            qt.Delete();
                                            ComUtilities.Release(ref qt);
                                            qt = null; // Prevent double-release in finally block

                                            // Recreate with new schema (query is still valid here)
                                            dynamic? newQt = CreateQueryTableForQuery(sheet, query);
                                            try
                                            {
                                                newQt.Refresh(false); // Synchronous refresh
                                                queryTableRecreated = true;
                                            }
                                            finally
                                            {
                                                ComUtilities.Release(ref newQt!);
                                            }
                                            break; // Only one QueryTable per query per sheet
                                        }
                                    }
                                    finally
                                    {
                                        ComUtilities.Release(ref qt);
                                    }
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref queryTables);
                                ComUtilities.Release(ref sheet);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref sheets);
                    }

                    // Release query and queries now that QueryTable recreation is done
                    ComUtilities.Release(ref query!);
                    ComUtilities.Release(ref queries!);

                    // Step 2: Refresh connection ONLY if no QueryTables were recreated
                    // (QueryTable refresh already happened above; connection refresh would interfere)
                    if (!queryTableRecreated)
                    {
                        try
                        {
                            RefreshConnectionByQueryName(ctx.Book, queryName);
                            result.Success = true;
                        }
                        catch (COMException comEx)
                        {
                            result.Success = false;
                            result.ErrorMessage = $"M code updated but refresh failed: {ParsePowerQueryError(comEx)}";
                        }
                    }
                    else
                    {
                        result.Success = true;
                    }

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
                    ComUtilities.Release(ref query!);
                    ComUtilities.Release(ref queries!);
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
    /// <param name="loadMode">Where to load the data</param>
    /// <param name="targetSheet">Target worksheet (required for LoadToTable/LoadToBoth)</param>
    /// <returns>Result with load configuration and refresh status</returns>
    public PowerQueryLoadResult LoadTo(
        IExcelBatch batch,
        string queryName,
        PowerQueryLoadMode loadMode,
        string? targetSheet = null)
    {
        var result = new PowerQueryLoadResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName,
            LoadDestination = loadMode,
            WorksheetName = targetSheet
        };

        try
        {
            if ((loadMode == PowerQueryLoadMode.LoadToTable || loadMode == PowerQueryLoadMode.LoadToBoth)
                && string.IsNullOrWhiteSpace(targetSheet))
            {
                result.Success = false;
                result.ErrorMessage = "Worksheet name required for LoadToTable/LoadToBoth";
                return result;
            }

            return batch.Execute((ctx, ct) =>
            {
                dynamic? queries = null;
                dynamic? query = null;
                dynamic? sheet = null;
                dynamic? queryTable = null;

                try
                {
                    queries = ctx.Book.Queries;
                    query = ComUtilities.FindQuery(ctx.Book, queryName);

                    if (query == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        return result;
                    }

                    // Apply load destination
                    switch (loadMode)
                    {
                        case PowerQueryLoadMode.LoadToTable:
                            // Check if sheet already exists - require explicit deletion by user
                            // Bug fix for Issue #170: Prevent silent failures when sheet exists
                            dynamic? worksheetsCheck = null;
                            try
                            {
                                worksheetsCheck = ctx.Book.Worksheets;
                                try
                                {
                                    dynamic? existingSheet = worksheetsCheck.Item(targetSheet!);
                                    if (existingSheet != null)
                                    {
                                        ComUtilities.Release(ref existingSheet!);
                                        result.Success = false;
                                        result.ErrorMessage = $"Cannot load query to sheet '{targetSheet}': worksheet already exists.";
                                        return result;
                                    }
                                }
                                catch (COMException)
                                {
                                    // Sheet doesn't exist - this is expected, continue
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref worksheetsCheck!);
                            }

                            // Create new sheet for query data
                            dynamic? worksheetsLoadToTable = null;
                            try
                            {
                                worksheetsLoadToTable = ctx.Book.Worksheets;
                                sheet = worksheetsLoadToTable.Add();
                                sheet.Name = targetSheet;
                            }
                            finally
                            {
                                ComUtilities.Release(ref worksheetsLoadToTable!);
                            }

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
                                ComUtilities.Release(ref dmConnectionLoadTo!);
                                ComUtilities.Release(ref connectionsLoadTo!);
                            }
                            break;

                        case PowerQueryLoadMode.LoadToBoth:
                            // Check if sheet already exists - require explicit deletion by user
                            // Bug fix for Issue #170: Prevent silent failures when sheet exists
                            dynamic? worksheetsCheckBoth = null;
                            try
                            {
                                worksheetsCheckBoth = ctx.Book.Worksheets;
                                try
                                {
                                    dynamic? existingSheet = worksheetsCheckBoth.Item(targetSheet!);
                                    if (existingSheet != null)
                                    {
                                        ComUtilities.Release(ref existingSheet!);
                                        result.Success = false;
                                        result.ErrorMessage = $"Cannot load query to sheet '{targetSheet}': worksheet already exists.";
                                        return result;
                                    }
                                }
                                catch (COMException)
                                {
                                    // Sheet doesn't exist - this is expected, continue
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref worksheetsCheckBoth!);
                            }

                            // Create new sheet for query data
                            dynamic? worksheetsLoadToBoth = null;
                            try
                            {
                                worksheetsLoadToBoth = ctx.Book.Worksheets;
                                sheet = worksheetsLoadToBoth.Add();
                                sheet.Name = targetSheet;
                            }
                            finally
                            {
                                ComUtilities.Release(ref worksheetsLoadToBoth!);
                            }

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
                                ComUtilities.Release(ref dmConnectionLoadToBoth!);
                                ComUtilities.Release(ref connectionsLoadToBoth!);
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
                    ComUtilities.Release(ref queryTable!);
                    ComUtilities.Release(ref sheet!);
                    ComUtilities.Release(ref query!);
                    ComUtilities.Release(ref queries!);
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
    public OperationResult Unload(
        IExcelBatch batch,
        string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "unload"
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? sheets = null;

            try
            {
                queries = ctx.Book.Queries;
                query = ComUtilities.FindQuery(ctx.Book, queryName);

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
                                ComUtilities.Release(ref qt!);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref queryTables!);
                        ComUtilities.Release(ref sheet!);
                    }
                }

                result.Success = true;

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
                ComUtilities.Release(ref sheets!);
                ComUtilities.Release(ref query!);
                ComUtilities.Release(ref queries!);
            }
        }, cancellationToken: default);
    }

    /// <summary>
    /// Helper method to create QueryTable for a query
    /// </summary>
    private static dynamic CreateQueryTableForQuery(dynamic sheet, dynamic query)
    {
        string queryName = query.Name;
        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";

        // Clear worksheet to prevent column accumulation (regression test requirement)
        dynamic? usedRange = null;
        try
        {
            usedRange = sheet.UsedRange;
            usedRange.Clear();
        }
        finally
        {
            ComUtilities.Release(ref usedRange);
        }

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
        queryTable.PreserveColumnInfo = false;  // Allow column structure changes when M code updates

        // Note: Caller is responsible for calling Refresh(false) after QueryTable is returned
        return queryTable;
    }
}

