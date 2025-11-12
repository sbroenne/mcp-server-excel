using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query operations - FilePath-based API using direct FileHandleManager integration
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public async Task<PowerQueryListResult> ListAsync(string filePath)
    {
        var result = new PowerQueryListResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? queriesCollection = null;
                try
                {
                    queriesCollection = handle.Workbook.Queries;
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
                                connections = handle.Workbook.Connections;
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
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error accessing Power Queries: {ex.Message}";

                    string extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension == ".xls")
                    {
                        result.ErrorMessage += " (.xls files don't support Power Query)";
                    }
                }
                finally
                {
                    ComUtilities.Release(ref queriesCollection);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> ViewAsync(string filePath, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = filePath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? query = null;
                try
                {
                    query = ComUtilities.FindQuery(handle.Workbook, queryName);
                    if (query == null)
                    {
                        var queryNames = GetQueryNames(handle.Workbook);
                        string? suggestion = FindClosestMatch(queryName, queryNames);

                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        if (suggestion != null)
                        {
                            result.ErrorMessage += $". Did you mean '{suggestion}'?";
                        }
                        return;
                    }

                    string mCode = query.Formula;
                    result.MCode = mCode;
                    result.CharacterCount = mCode.Length;

                    // Check if connection only
                    bool isConnectionOnly = true;
                    dynamic? connections = null;
                    try
                    {
                        connections = handle.Workbook.Connections;
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
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error viewing query: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> ExportAsync(string filePath, string queryName, string outputFile)
    {
        var result = new OperationResult { FilePath = filePath };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? query = null;
                try
                {
                    query = ComUtilities.FindQuery(handle.Workbook, queryName);
                    if (query == null)
                    {
                        var queryNames = GetQueryNames(handle.Workbook);
                        string? suggestion = FindClosestMatch(queryName, queryNames);

                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        if (suggestion != null)
                        {
                            result.ErrorMessage += $". Did you mean '{suggestion}'?";
                        }
                        return;
                    }

                    string mCode = query.Formula;
                    File.WriteAllText(outputFile, mCode);

                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error exporting query: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<PowerQueryRefreshResult> RefreshAsync(string filePath, string queryName)
    {
        return await RefreshAsync(filePath, queryName, timeout: null);
    }

    /// <inheritdoc />
    public async Task<PowerQueryRefreshResult> RefreshAsync(string filePath, string queryName, TimeSpan? timeout)
    {
        var result = new PowerQueryRefreshResult
        {
            FilePath = filePath,
            QueryName = queryName,
            RefreshTime = DateTime.Now
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            var timeoutValue = timeout ?? TimeSpan.FromMinutes(5);
            var cts = new CancellationTokenSource(timeoutValue);

            await Task.Run(() =>
            {
                dynamic? query = null;
                try
                {
                    query = ComUtilities.FindQuery(handle.Workbook, queryName);
                    if (query == null)
                    {
                        var queryNames = GetQueryNames(handle.Workbook);
                        string? suggestion = FindClosestMatch(queryName, queryNames);

                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        if (suggestion != null)
                        {
                            result.ErrorMessage += $". Did you mean '{suggestion}'?";
                        }
                        return;
                    }

                    // Check if query has a connection to refresh
                    try
                    {
                        // Use RefreshConnectionByQueryName helper to avoid code duplication
                        RefreshConnectionByQueryName(handle.Workbook, queryName);

                        // Check for errors after refresh
                        result.HasErrors = false;
                        result.Success = true;
                        result.LoadedToSheet = DetermineLoadedSheet(handle.Workbook, queryName);

                        // Determine if connection-only based on whether it's loaded to a sheet OR Data Model
                        bool isLoadedToDataModel = IsQueryLoadedToDataModel(handle.Workbook, queryName);
                        result.IsConnectionOnly = string.IsNullOrEmpty(result.LoadedToSheet) && !isLoadedToDataModel;
                    }
                    catch (COMException comEx)
                    {
                        // Capture detailed error information
                        result.Success = false;
                        result.HasErrors = true;
                        result.ErrorMessages.Add(ParsePowerQueryError(comEx));
                        result.ErrorMessage = string.Join("; ", result.ErrorMessages);

                        var errorCategory = CategorizeError(comEx);
                    }

                    // If no connection found, check if query is loaded to worksheet or data model
                    if (!result.Success && result.ErrorMessages.Count == 0)
                    {
                        ComUtilities.Release(ref query);

                        // Check if there are QueryTables that reference this query OR if it's in Data Model
                        string? loadedSheet = DetermineLoadedSheet(handle.Workbook, queryName);
                        bool isLoadedToDataModel = IsQueryLoadedToDataModel(handle.Workbook, queryName);

                        if (loadedSheet != null || isLoadedToDataModel)
                        {
                            // Query is loaded to a worksheet via QueryTable or Data Model
                            result.Success = true;
                            result.IsConnectionOnly = false;
                            result.LoadedToSheet = loadedSheet;
                        }
                        else
                        {
                            // Truly connection-only (no connection, no QueryTables)
                            result.Success = true;
                            result.IsConnectionOnly = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error refreshing query: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }, cts.Token);

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadConfigResult> GetLoadConfigAsync(string filePath, string queryName)
    {
        var result = new PowerQueryLoadConfigResult
        {
            FilePath = filePath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? query = null;
                dynamic? worksheets = null;
                dynamic? connections = null;
                dynamic? names = null;
                try
                {
                    query = ComUtilities.FindQuery(handle.Workbook, queryName);
                    if (query == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        return;
                    }

                    // Check for QueryTables first (table loading)
                    bool hasTableConnection = false;
                    bool hasDataModelConnection = false;
                    string? targetSheet = null;

                    worksheets = handle.Workbook.Worksheets;
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

                    // Check for Data Model connection
                    connections = handle.Workbook.Connections;
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
                                hasDataModelConnection = true;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }

                    // Determine load mode
                    PowerQueryLoadMode loadMode;
                    if (hasTableConnection && hasDataModelConnection)
                    {
                        loadMode = PowerQueryLoadMode.LoadToBoth;
                    }
                    else if (hasTableConnection)
                    {
                        loadMode = PowerQueryLoadMode.LoadToTable;
                    }
                    else if (hasDataModelConnection)
                    {
                        loadMode = PowerQueryLoadMode.LoadToDataModel;
                    }
                    else
                    {
                        loadMode = PowerQueryLoadMode.ConnectionOnly;
                    }

                    result.LoadMode = loadMode;
                    result.TargetSheet = targetSheet;
                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error getting load config: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref names);
                    ComUtilities.Release(ref connections);
                    ComUtilities.Release(ref worksheets);
                    ComUtilities.Release(ref query);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(string filePath, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "pq-delete"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? query = null;
                dynamic? queriesCollection = null;
                try
                {
                    query = ComUtilities.FindQuery(handle.Workbook, queryName);
                    if (query == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        return;
                    }

                    queriesCollection = handle.Workbook.Queries;
                    queriesCollection.Item(queryName).Delete();

                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error deleting query: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref queriesCollection);
                    ComUtilities.Release(ref query);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<WorksheetListResult> ListExcelSourcesAsync(string filePath)
    {
        var result = new WorksheetListResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? worksheets = null;
                try
                {
                    worksheets = handle.Workbook.Worksheets;
                    int count = worksheets.Count;

                    for (int i = 1; i <= count; i++)
                    {
                        dynamic? sheet = null;
                        try
                        {
                            sheet = worksheets.Item(i);
                            string sheetName = sheet.Name ?? $"Sheet{i}";
                            result.Worksheets.Add(new WorksheetInfo { Name = sheetName });
                        }
                        finally
                        {
                            ComUtilities.Release(ref sheet);
                        }
                    }

                    result.Success = true;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error listing worksheets: {ex.Message}";
                }
                finally
                {
                    ComUtilities.Release(ref worksheets);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }

    /// <inheritdoc />
    public async Task<PowerQueryCreateResult> CreateAsync(string filePath, string queryName, string mCodeFile, PowerQueryLoadMode loadMode = PowerQueryLoadMode.LoadToTable, string? targetSheet = null)
    {
        // TODO: Implement direct FileHandleManager pattern
        // Complex operation requiring QueryTable creation, loading, etc.
        // For now, use batch-based method via temporary session
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await CreateAsync(batch, queryName, mCodeFile, loadMode, targetSheet);
        // Note: No automatic save - caller must save explicitly via WorkbookCommands.SaveAsync
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateAsync(string filePath, string queryName, string mCodeFile)
    {
        // TODO: Implement direct FileHandleManager pattern
        // Complex operation requiring formula update
        // For now, use batch-based method via temporary session
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await UpdateAsync(batch, queryName, mCodeFile);
        // Note: No automatic save - caller must save explicitly via WorkbookCommands.SaveAsync
        return result;
    }

    /// <inheritdoc />
    public async Task<PowerQueryLoadResult> LoadToAsync(string filePath, string queryName, PowerQueryLoadMode loadMode, string? targetSheet = null)
    {
        // TODO: Implement direct FileHandleManager pattern
        // Complex operation requiring QueryTable/connection management
        // For now, use batch-based method via temporary session
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await LoadToAsync(batch, queryName, loadMode, targetSheet);
        // Note: No automatic save - caller must save explicitly via WorkbookCommands.SaveAsync
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> UnloadAsync(string filePath, string queryName)
    {
        // TODO: Implement direct FileHandleManager pattern
        // Complex operation requiring QueryTable/connection removal
        // For now, use batch-based method via temporary session
        await using var batch = await ExcelSession.BeginBatchAsync(filePath);
        var result = await UnloadAsync(batch, queryName);
        // Note: No automatic save - caller must save explicitly via WorkbookCommands.SaveAsync
        return result;
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAllAsync(string filePath)
    {
        var result = new OperationResult
        {
            FilePath = filePath
        };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? queries = null;

                try
                {
                    queries = handle.Workbook.Queries;
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
                            var connection = FindConnectionForQuery(handle.Workbook, queryName);
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
                            ComUtilities.Release(ref query!);
                        }
                    }

                    // ✅ Rule 0: Success = false when errors exist
                    if (errors.Count > 0)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Some queries failed to refresh: {string.Join(", ", errors)}";
                    }
                    else
                    {
                        result.Success = true;
                    }
                }
                catch (COMException ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Excel COM error refreshing queries: {ex.Message}";
                    result.IsRetryable = ex.HResult == -2147417851;
                }
                finally
                {
                    ComUtilities.Release(ref queries!);
                }
            });

            return result;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Failed to access workbook: {ex.Message}";
            return result;
        }
    }
}
