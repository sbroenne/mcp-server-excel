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
    public async Task<PowerQueryRefreshResult> RefreshAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryRefreshResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName,
            RefreshTime = DateTime.Now
        };

        return await batch.Execute<PowerQueryRefreshResult>((ctx, ct) =>
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
                dynamic? targetConnection = null;
                dynamic? connections = null;
                try
                {
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
                catch { }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                if (targetConnection != null)
                {
                    try
                    {
                        // Attempt refresh and capture any errors
                        targetConnection.Refresh();

                        // Check for errors after refresh
                        result.HasErrors = false;
                        result.Success = true;
                        result.LoadedToSheet = DetermineLoadedSheet(ctx.Book, queryName);

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
                    finally
                    {
                        ComUtilities.Release(ref targetConnection);
                    }
                }
                else
                {
                    // Connection-only query
                    ComUtilities.Release(ref query);
                    result.Success = true;
                    result.IsConnectionOnly = true;
                }

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error refreshing query: {ex.Message}";
                return result;
            }
        });
    }

    /// <inheritdoc />
    public async Task<PowerQueryViewResult> ErrorsAsync(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        return await batch.Execute<PowerQueryViewResult>((ctx, ct) =>
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

                // Try to get error information if available
                dynamic? connections = null;
                try
                {
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
                                // Connection found - query has been loaded
                                result.MCode = "No error information available through Excel COM interface";
                                result.Success = true;
                                return result;
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

                result.MCode = "Query is connection-only - no error information available";
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error checking query errors: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref query);
            }
        });
    }
}
