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
    public PowerQueryRefreshResult Refresh(IExcelBatch batch, string queryName, TimeSpan timeout, IProgress<ProgressInfo>? progress = null)
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

        timeout = NormalizeRefreshTimeout(timeout);

        using var timeoutCts = new CancellationTokenSource(timeout);
        string? queryFormula = null;

        try
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.WorkbookQuery? query = null;
                try
                {
                    query = PowerQuery.PowerQueryHelpers.FindQueryByExactName(ctx.Book, queryName);
                    if (query == null)
                    {
                        throw new OperationFailureException(
                            OperationFailureCategory.NotFound,
                            $"Query '{queryName}' not found.");
                    }

                    queryFormula = query.Formula?.ToString();

                    // Refresh the query - exceptions propagate from both:
                    // - QueryTable.Refresh() for worksheet queries
                    // - Connection.Refresh() for Data Model queries
                    progress?.Report(new ProgressInfo { Current = 0, Total = 1, Message = $"Refreshing '{queryName}'" });
                    bool refreshed;
                    try
                    {
                        refreshed = RefreshConnectionByQueryName(ctx.Book, queryName, timeoutCts.Token);
                    }
                    catch (Exception ex) when (TryWrapPowerQueryException(ex, out var pqEx))
                    {
                        throw pqEx!;
                    }

                    if (!refreshed)
                    {
                        throw new OperationFailureException(
                            OperationFailureCategory.Prerequisite,
                            MissingRefreshDestinationMessage(queryName));
                    }

                    result.HasErrors = false;
                    result.Success = true;
                    var loadState = DetectLoadState(ctx.Book, queryName, ct);
                    result.LoadedToSheet = loadState.TargetSheet;
                    result.IsConnectionOnly = loadState.IsConnectionOnly;

                    progress?.Report(new ProgressInfo { Current = 1, Total = 1, Message = $"Refreshed '{queryName}'" });
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }, timeoutCts.Token);
        }
        catch (TimeoutException ex) when (IsLikelyPrivacyFirewallRisk(queryFormula))
        {
            throw new PowerQueryCommandException(
                $"Likely Formula.Firewall/privacy-blocked refresh for query '{queryName}'. The query combines Excel.CurrentWorkbook with an external data source and Excel may have shown a privacy/modal prompt instead of returning a normal refresh error.",
                "Privacy",
                ex);
        }
    }

    /// <summary>
    /// Refreshes every Power Query in the workbook. Queries with nothing to refresh on their
    /// own (parameter and connection-only staging queries) are skipped. A failed query is
    /// recorded and the remaining queries are still refreshed; nothing is rolled back.
    /// Timeouts, cancellation, and Excel disconnects stop the operation immediately.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="timeout">Maximum time to wait for all refreshes to complete</param>
    /// <param name="progress">Optional progress reporter</param>
    public PowerQueryRefreshAllResult RefreshAll(IExcelBatch batch, TimeSpan timeout = default, IProgress<ProgressInfo>? progress = null)
    {
        timeout = NormalizeRefreshTimeout(timeout);

        using var timeoutCts = new CancellationTokenSource(timeout);

        return batch.Execute((ctx, ct) =>
        {
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token);
            var queryNames = ReadQueryNames(ctx.Book, linkedCts.Token);

            // Use the same robust strategy as single-query Refresh:
            // 1) QueryTable.Refresh(false) for worksheet-loaded queries
            // 2) Connection.Refresh() for Data Model queries
            var result = RefreshQueries(
                queryNames,
                queryName => RefreshConnectionByQueryName(ctx.Book, queryName, linkedCts.Token),
                linkedCts.Token,
                progress);
            result.FilePath = batch.WorkbookPath;
            return result;
        }, timeoutCts.Token);
    }

    internal static PowerQueryRefreshAllResult RefreshQueries(
        IReadOnlyList<string> queryNames,
        Func<string, bool> refreshQuery,
        CancellationToken cancellationToken,
        IProgress<ProgressInfo>? progress = null)
    {
        var result = new PowerQueryRefreshAllResult();
        int total = queryNames.Count;

        for (int i = 0; i < total; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string queryName = queryNames[i];
            progress?.Report(new ProgressInfo { Current = i, Total = total, Message = $"Refreshing '{queryName}' ({i + 1}/{total})" });

            try
            {
                if (refreshQuery(queryName))
                {
                    result.RefreshedQueries.Add(queryName);
                }
                else
                {
                    result.SkippedQueries.Add(new PowerQueryRefreshSkip
                    {
                        QueryName = queryName,
                        Reason = NothingToRefreshReason
                    });
                }
            }
            catch (Exception ex) when (IsRecordableQueryFailure(ex, cancellationToken))
            {
                result.FailedQueries.Add(CreateRefreshFailure(queryName, ex));
            }
        }

        progress?.Report(new ProgressInfo { Current = total, Total = total, Message = "Refresh-all finished" });

        string summary = SummarizeRefreshAll(result);
        if (result.FailedQueries.Count > 0)
        {
            result.Success = false;
            result.ErrorMessage =
                $"{result.FailedQueries.Count} of {total} queries failed to refresh: " +
                $"{string.Join(", ", result.FailedQueries.Select(f => $"'{f.QueryName}'"))}. " +
                "Nothing was rolled back; " + summary +
                " Inspect failedQueries for each error, then refresh failed queries by name after fixing them.";
            return result;
        }

        result.Success = true;
        result.Message = summary;
        return result;
    }

    private const string NothingToRefreshReason =
        "No worksheet table, Data Model table, or workbook connection to refresh. " +
        "Parameter and connection-only staging queries are evaluated when the loaded queries that use them refresh.";

    private static List<string> ReadQueryNames(Excel.Workbook workbook, CancellationToken cancellationToken)
    {
        var names = new List<string>();
        Excel.Queries? queries = null;
        try
        {
            queries = workbook.Queries;
            int count = queries.Count;
            for (int i = 1; i <= count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.WorkbookQuery? query = null;
                try
                {
                    query = queries.Item(i);
                    names.Add(query.Name);
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref queries);
        }

        return names;
    }

    private static bool IsRecordableQueryFailure(Exception exception, CancellationToken cancellationToken)
    {
        // Timeouts, cancellation, and a dead/disconnected Excel must stop the whole operation
        // so the batch and Service layers can report them and recover the session.
        if (cancellationToken.IsCancellationRequested ||
            exception is OperationCanceledException or TimeoutException)
        {
            return false;
        }

        for (Exception? current = exception; current != null; current = current.InnerException)
        {
            if (current is COMException comException && IsFatalExcelDisconnect(comException))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsFatalExcelDisconnect(COMException exception) =>
        exception.HResult is ResiliencePipelines.RPC_E_DISCONNECTED
            or ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE
            or ResiliencePipelines.RPC_E_CALL_FAILED;

    private static PowerQueryRefreshFailure CreateRefreshFailure(string queryName, Exception exception)
    {
        string? category = exception switch
        {
            PowerQueryCommandException pq => pq.ErrorCategory,
            OperationFailureException failure => failure.ErrorCategory.ToString(),
            _ => ClassifyPowerQueryError(exception.Message)
        };

        // Our own wrapper exceptions carry no HRESULT; report the underlying Excel/COM one.
        Exception? origin = exception is PowerQueryCommandException or OperationFailureException
            ? exception.InnerException
            : exception;

        return new PowerQueryRefreshFailure
        {
            QueryName = queryName,
            ErrorCategory = category,
            ErrorMessage = exception.Message,
            ExceptionType = exception.GetType().Name,
            HResult = origin != null ? $"0x{origin.HResult:X8}" : null
        };
    }

    private static string SummarizeRefreshAll(PowerQueryRefreshAllResult result)
    {
        string refreshed = result.RefreshedQueries.Count == 0
            ? "no queries were refreshed."
            : $"refreshed {result.RefreshedQueries.Count}: {string.Join(", ", result.RefreshedQueries.Select(n => $"'{n}'"))}.";
        string skipped = result.SkippedQueries.Count == 0
            ? string.Empty
            : $" Skipped {result.SkippedQueries.Count} with nothing to refresh: {string.Join(", ", result.SkippedQueries.Select(s => $"'{s.QueryName}'"))}.";
        return char.ToUpperInvariant(refreshed[0]) + refreshed[1..] + skipped;
    }

    internal static TimeSpan NormalizeRefreshTimeout(TimeSpan timeout)
    {
        if (timeout <= TimeSpan.Zero)
        {
            return ComInteropConstants.DataOperationTimeout;
        }

        TimeSpan maximum = TimeSpan.FromMilliseconds(uint.MaxValue - 1);
        return timeout > maximum ? maximum : timeout;
    }

    private static string MissingRefreshDestinationMessage(string queryName) =>
        $"Could not find connection or table for query '{queryName}'. " +
        "For definition-only staging queries, refresh the loaded dependent queries by name " +
        "using powerquery refresh (MCP: action='refresh', query_name; CLI: --query-name). " +
        "Inspect powerquery get-load-config first, then refresh dependent PivotTables.";

}
