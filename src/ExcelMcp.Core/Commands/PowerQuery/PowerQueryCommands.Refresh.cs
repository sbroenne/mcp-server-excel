using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

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
            timeout = TimeSpan.FromMinutes(5); // Default timeout when not specified
        }

        using var timeoutCts = new CancellationTokenSource(timeout);

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookQuery? query = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // Refresh the query - exceptions propagate from both:
                // - QueryTable.Refresh() for worksheet queries
                // - Connection.Refresh() for Data Model queries
                bool refreshed = RefreshConnectionByQueryName(ctx.Book, queryName);

                if (!refreshed)
                {
                    throw new InvalidOperationException($"Could not find connection or table for query '{queryName}'.");
                }

                result.HasErrors = false;
                result.Success = true;
                result.LoadedToSheet = DetermineLoadedSheet(ctx.Book, queryName);

                bool isLoadedToDataModel = IsQueryLoadedToDataModel(ctx.Book, queryName);
                result.IsConnectionOnly = string.IsNullOrEmpty(result.LoadedToSheet) && !isLoadedToDataModel;

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
    public OperationResult RefreshAll(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Queries? queries = null;

            try
            {
                queries = ctx.Book.Queries;
                int totalQueries = queries.Count;
                var errors = new List<string>();

                for (int i = 1; i <= totalQueries; i++)
                {
                    Excel.WorkbookQuery? query = null;
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

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
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



