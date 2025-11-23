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
    public PowerQueryRefreshResult Refresh(IExcelBatch batch, string queryName, TimeSpan timeout)
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
            throw new ArgumentException(validationError, nameof(queryName));
        }

        if (timeout <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(timeout), "Timeout must be greater than zero.");
        }

        using var timeoutCts = new CancellationTokenSource(timeout);

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

                    string errorMsg = $"Query '{queryName}' not found";
                    if (suggestion != null)
                    {
                        errorMsg += $". Did you mean '{suggestion}'?";
                    }
                    throw new InvalidOperationException(errorMsg);
                }

                try
                {
                    RefreshConnectionByQueryName(ctx.Book, queryName);

                    result.HasErrors = false;
                    result.Success = true;
                    result.LoadedToSheet = DetermineLoadedSheet(ctx.Book, queryName);

                    bool isLoadedToDataModel = IsQueryLoadedToDataModel(ctx.Book, queryName);
                    result.IsConnectionOnly = string.IsNullOrEmpty(result.LoadedToSheet) && !isLoadedToDataModel;
                }
                catch (COMException comEx)
                {
                    result.Success = false;
                    result.HasErrors = true;
                    result.ErrorMessages.Add(ParsePowerQueryError(comEx));
                    result.ErrorMessage = string.Join("; ", result.ErrorMessages);
                }

                if (!result.Success && result.ErrorMessages.Count == 0)
                {
                    ComUtilities.Release(ref query);
                    query = null;

                    string? loadedSheet = DetermineLoadedSheet(ctx.Book, queryName);
                    bool isLoadedToDataModel = IsQueryLoadedToDataModel(ctx.Book, queryName);

                    if (loadedSheet != null || isLoadedToDataModel)
                    {
                        result.Success = true;
                        result.IsConnectionOnly = false;
                        result.LoadedToSheet = loadedSheet;
                    }
                    else
                    {
                        result.Success = true;
                        result.IsConnectionOnly = true;
                    }
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        }, timeoutCts.Token);
    }

    /// <summary>
    /// Refreshes all Power Query queries in the workbook
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <exception cref="InvalidOperationException">Thrown when refresh fails</exception>
    public void RefreshAll(IExcelBatch batch)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;

            try
            {
                queries = ctx.Book.Queries;
                int totalQueries = queries.Count;
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

                // Throw if any errors occurred
                if (errors.Count > 0)
                {
                    throw new InvalidOperationException($"Some queries failed to refresh: {string.Join(", ", errors)}");
                }

                return 0;
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

