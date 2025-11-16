using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query refresh operations
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public PowerQueryRefreshResult Refresh(IExcelBatch batch, string queryName)
    {
        return Refresh(batch, queryName, timeout: null);
    }

    /// <inheritdoc />
    public PowerQueryRefreshResult Refresh(IExcelBatch batch, string queryName, TimeSpan? timeout)
    {
        var result = new PowerQueryRefreshResult
        {
            FilePath = batch.WorkbookPath,
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

                // Check if query has a connection to refresh
                try
                {
                    // Use RefreshConnectionByQueryName helper to avoid code duplication
                    RefreshConnectionByQueryName(ctx.Book, queryName);

                    // Check for errors after refresh
                    result.HasErrors = false;
                    result.Success = true;
                    result.LoadedToSheet = DetermineLoadedSheet(ctx.Book, queryName);

                    // Determine if connection-only based on whether it's loaded to a sheet OR Data Model
                    bool isLoadedToDataModel = IsQueryLoadedToDataModel(ctx.Book, queryName);
                    result.IsConnectionOnly = string.IsNullOrEmpty(result.LoadedToSheet) && !isLoadedToDataModel;

                    // Add workflow guidance
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
                    string? loadedSheet = DetermineLoadedSheet(ctx.Book, queryName);
                    bool isLoadedToDataModel = IsQueryLoadedToDataModel(ctx.Book, queryName);

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

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error refreshing query: {ex.Message}";
                return result;
            }
        });  // Default 5 minutes for Power Query refresh, LLM can override
    }

    /// <summary>
    /// Refreshes all Power Query queries in the workbook
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <returns>Operation result with refresh summary</returns>
    public OperationResult RefreshAll(IExcelBatch batch)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath
        };

        return batch.Execute((ctx, ct) =>
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
                ComUtilities.Release(ref queries!);
            }
        }, cancellationToken: default);
    }

    /// <summary>
    /// Helper method to find connection for a query
    /// </summary>
    private static dynamic? FindConnectionForQuery(dynamic workbook, string queryName)
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
                        ComUtilities.Release(ref conn!);
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref connections!);
        }

        return null;
    }
}

