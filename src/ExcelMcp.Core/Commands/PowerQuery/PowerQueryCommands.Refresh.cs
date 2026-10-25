using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PowerQueryTrace;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query refresh operations
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public PowerQueryRefreshResult Refresh(
        IExcelBatch batch,
        string queryName,
        TimeSpan timeout,
        bool trace = false,
        IProgress<ProgressInfo>? progress = null)
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
        DateTime traceStartUtc = DateTime.UtcNow;
        var started = System.Diagnostics.Stopwatch.GetTimestamp();

        string? queryFormula = null;

        try
        {
            var refreshResult = ExcelBatchWait.Execute(batch, timeout, (ctx, ct) =>
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
                        refreshed = RefreshConnectionByQueryName(ctx.Book, queryName, ct);
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
                    ReadLoadMetrics(ctx.Book, queryName, result, ct);
                    result.DurationMs = System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;

                    progress?.Report(new ProgressInfo { Current = 1, Total = 1, Message = $"Refreshed '{queryName}'" });
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            });
            if (trace)
            {
                AddRefreshTrace(refreshResult, batch.ExcelProcessId, traceStartUtc, DateTime.UtcNow, failed: false);
            }

            return refreshResult;
        }
        catch (TimeoutException ex) when (
            ex.InnerException is not OperationCanceledException && IsLikelyPrivacyFirewallRisk(queryFormula))
        {
            const string category = "PrivacyFirewall";
            throw new PowerQueryCommandException(
                AddPowerQueryRemediation(
                    $"Likely Formula.Firewall/privacy-blocked refresh for query '{queryName}'. The query combines Excel.CurrentWorkbook with an external data source and Excel may have shown a privacy/modal prompt instead of returning a normal refresh error.",
                    category),
                category,
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
    public PowerQueryRefreshAllResult RefreshAll(
        IExcelBatch batch,
        TimeSpan timeout = default,
        bool refreshPivotTables = true,
        bool trace = false,
        IProgress<ProgressInfo>? progress = null)
    {
        timeout = NormalizeRefreshTimeout(timeout);
        DateTime traceStartUtc = DateTime.UtcNow;

        var result = ExcelBatchWait.Execute(batch, timeout, (ctx, ct) =>
        {
            var queryNames = ReadQueryNames(ctx.Book, ct);
            var queryMetrics = new List<PowerQueryOperationResult>();

            // Use the same robust strategy as single-query Refresh:
            // 1) QueryTable.Refresh(false) for worksheet-loaded queries
            // 2) Connection.Refresh() for Data Model queries
            var result = RefreshQueries(
                queryNames,
                queryName =>
                {
                    var started = System.Diagnostics.Stopwatch.GetTimestamp();
                    if (!RefreshConnectionByQueryName(ctx.Book, queryName, ct))
                        return false;
                    var metrics = new PowerQueryOperationResult { Success = true, QueryName = queryName };
                    ReadLoadMetrics(ctx.Book, queryName, metrics, ct);
                    metrics.DurationMs = System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;
                    queryMetrics.Add(metrics);
                    return true;
                },
                ct,
                progress);
            if (refreshPivotTables && result.RefreshedQueries.Count > 0)
            {
                RefreshDependentPivotTables(ctx.Book, result, ct, progress);
            }

            FinalizeRefreshAllResult(result);
            result.QueryMetrics = queryMetrics;
            result.FilePath = batch.WorkbookPath;
            return result;
        });
        if (trace)
        {
            AddRefreshTrace(result, batch.ExcelProcessId, traceStartUtc, DateTime.UtcNow, failed: !result.Success);
        }

        return result;
    }

    internal void AddRefreshTrace(
        PowerQueryRefreshResult result,
        int? processId,
        DateTime startUtc,
        DateTime endUtc,
        bool failed)
    {
        var summary = ReadRefreshTrace(processId, startUtc, endUtc);
        result.TraceSummary = summary;
        result.TraceNote = CreateRefreshTraceNote(summary, endUtc - startUtc, failed);
    }

    internal void AddRefreshTrace(
        PowerQueryRefreshAllResult result,
        int? processId,
        DateTime startUtc,
        DateTime endUtc,
        bool failed)
    {
        var summary = ReadRefreshTrace(processId, startUtc, endUtc);
        result.TraceSummary = summary;
        result.TraceNote = CreateRefreshTraceNote(summary, endUtc - startUtc, failed);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Optional trace-file I/O failures are reported in traceSummary with Success=false and the original error message plus a traceNote; the completed refresh outcome is unchanged. Cancellation and unexpected failures propagate.")]
    private PowerQueryTraceSummaryResult ReadRefreshTrace(int? processId, DateTime startUtc, DateTime endUtc)
    {
        try
        {
            return _traceReader.Summarize(new PowerQueryTraceFilter(
                ProcessId: processId,
                SinceUtc: startUtc,
                UntilUtc: endUtc));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
        {
            return new PowerQueryTraceSummaryResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                ProcessId = processId,
                WindowStartUtc = startUtc,
                WindowEndUtc = endUtc
            };
        }
    }

    private static string? CreateRefreshTraceNote(
        PowerQueryTraceSummaryResult summary,
        TimeSpan elapsed,
        bool failed)
    {
        if (!summary.Success)
        {
            return "The refresh result above is unchanged, but the trace files could not be read. See traceSummary.errorMessage; do not repeat the refresh just to collect traces.";
        }

        var notes = new List<string>();
        if (summary.EntriesMatched == 0)
        {
            notes.Add(
                "No trace entries were found. Power Query reads tracing when Excel starts. Enable it with `excelcli settings set-pq-tracing --enabled true` (MCP: `settings`, action `set-pq-tracing` with `enabled=true`), restart Excel, then refresh again. This setting applies to the current Windows user and is not changed when a workbook or session closes. The refresh did not restart Excel.");
        }

        if (elapsed >= TimeSpan.FromSeconds(60))
        {
            notes.Add("This refresh took at least 60 seconds; use the trace summary to inspect data-source time and quiet gaps.");
        }

        if (failed)
        {
            notes.Add("One or more refresh operations failed; review the failed query list together with the trace summary.");
        }

        return notes.Count == 0 ? null : string.Join(" ", notes);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Transports recordable refresh failures in FailedQueries with native diagnostics; overall success is false, and cancellation/disconnection are excluded.")]
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

        progress?.Report(new ProgressInfo { Current = total, Total = total, Message = "Power Query refreshes finished" });

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

    internal static List<string> ReadQueryNames(Excel.Workbook workbook, CancellationToken cancellationToken)
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

    internal static bool IsRecordableQueryFailure(Exception exception, CancellationToken cancellationToken)
    {
        // Timeouts, cancellation (by the caller or by Excel), and a dead/disconnected Excel
        // must stop the whole operation so the batch and Service layers can report them and
        // recover the session.
        if (cancellationToken.IsCancellationRequested ||
            exception is OperationCanceledException or TimeoutException)
        {
            return false;
        }

        var pending = new Stack<Exception>();
        pending.Push(exception);
        while (pending.TryPop(out var current))
        {
            if (current is COMException comException && IsFatalExcelDisconnect(comException))
            {
                return false;
            }

            if (current is OperationFailureException { ErrorCategory: OperationFailureCategory.Cancelled })
            {
                return false;
            }

            if (current is AggregateException aggregate)
            {
                foreach (var innerException in aggregate.InnerExceptions)
                {
                    pending.Push(innerException);
                }
            }
            else if (current.InnerException is { } innerException)
            {
                pending.Push(innerException);
            }
        }

        return true;
    }

    private static bool IsFatalExcelDisconnect(COMException exception) =>
        exception.HResult is ResiliencePipelines.RPC_E_DISCONNECTED
            or ResiliencePipelines.RPC_S_SERVER_UNAVAILABLE
            or ResiliencePipelines.RPC_E_CALL_FAILED;

    internal static PowerQueryRefreshFailure CreateRefreshFailure(string queryName, Exception exception)
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
            ErrorMessage = AddPowerQueryRemediation(exception.Message, category),
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
        string pivotTables = result.RefreshedPivotTables.Count == 0
            ? string.Empty
            : $" Refreshed {result.RefreshedPivotTables.Count} dependent PivotTable(s): " +
              $"{string.Join(", ", result.RefreshedPivotTables.Select(n => $"'{n}'"))}.";
        return char.ToUpperInvariant(refreshed[0]) + refreshed[1..] + skipped + pivotTables;
    }

    private static void FinalizeRefreshAllResult(PowerQueryRefreshAllResult result)
    {
        string summary = SummarizeRefreshAll(result);
        result.Message = summary;
        if (result.FailedQueries.Count == 0 && result.FailedPivotTables.Count == 0)
        {
            result.Success = true;
            result.ErrorMessage = null;
            return;
        }

        result.Success = false;
        var errors = new List<string>();
        if (result.FailedQueries.Count > 0)
        {
            errors.Add(
                $"{result.FailedQueries.Count} query refresh(es) failed: " +
                $"{string.Join(", ", result.FailedQueries.Select(f => $"'{f.QueryName}'"))}");
        }

        if (result.FailedPivotTables.Count > 0)
        {
            errors.Add(
                $"{result.FailedPivotTables.Count} PivotTable refresh(es) failed: " +
                $"{string.Join(", ", result.FailedPivotTables.Select(f => $"'{f.PivotTableName}'"))}");
        }

        result.ErrorMessage =
            string.Join(". ", errors) + ". Nothing was rolled back. " + summary +
            " Inspect failedQueries and failedPivotTables for details.";
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Transports cache-classification and refresh errors as FailedPivotTables with original diagnostics; cancellation and session disconnection propagate.")]
    private static void RefreshDependentPivotTables(
        Excel.Workbook workbook,
        PowerQueryRefreshAllResult result,
        CancellationToken cancellationToken,
        IProgress<ProgressInfo>? progress)
    {
        var queryNames = new HashSet<string>(result.RefreshedQueries, StringComparer.OrdinalIgnoreCase);
        var (loadedTableNames, loadedTableRanges) =
            ReadWorksheetTableSourcesForQueries(workbook, queryNames, cancellationToken);
        bool refreshDataModelPivots =
            HasDataModelLoadForQueries(workbook, queryNames, cancellationToken);
        if (loadedTableNames.Count == 0 && loadedTableRanges.Count == 0 && !refreshDataModelPivots)
        {
            return;
        }

        var cacheOutcomes = new Dictionary<int, PowerQueryPivotRefreshFailure?>();
        Excel.Sheets? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.Worksheet? worksheet = null;
                Excel.PivotTables? pivotTables = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.Item[i];
                    string sheetName = worksheet.Name;
                    pivotTables = worksheet.PivotTables();
                    for (int j = 1; j <= pivotTables.Count; j++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        Excel.PivotTable? pivotTable = null;
                        Excel.PivotCache? pivotCache = null;
                        try
                        {
                            pivotTable = pivotTables.Item(j);
                            pivotCache = pivotTable.PivotCache();
                            bool isOlap = pivotCache.OLAP;
                            string pivotName = $"{sheetName}!{pivotTable.Name}";
                            int cacheIndex = pivotCache.Index;
                            bool isDependent;
                            try
                            {
                                isDependent = isOlap
                                    ? refreshDataModelPivots && IsDataModelPivotCache(pivotCache)
                                    : pivotCache.SourceType == Excel.XlPivotTableSourceType.xlDatabase &&
                                      IsPivotCacheBackedByQueryTable(
                                          workbook,
                                          pivotCache,
                                          loadedTableNames,
                                          loadedTableRanges);
                            }
                            catch (Exception ex) when (IsRecordableQueryFailure(ex, cancellationToken))
                            {
                                var classificationFailure = CreatePivotRefreshFailure(pivotName, ex);
                                cacheOutcomes.TryAdd(cacheIndex, classificationFailure);
                                result.FailedPivotTables.Add(classificationFailure);
                                continue;
                            }

                            if (!isDependent)
                            {
                                continue;
                            }

                            if (!cacheOutcomes.TryGetValue(cacheIndex, out var failure))
                            {
                                progress?.Report(new ProgressInfo
                                {
                                    Current = result.RefreshedPivotTables.Count + result.FailedPivotTables.Count,
                                    Total = result.RefreshedPivotTables.Count + result.FailedPivotTables.Count + 1,
                                    Message = $"Refreshing dependent PivotTable '{pivotName}'"
                                });
                                try
                                {
                                    PivotTableRefreshHelper.RefreshSynchronously(
                                        pivotTable,
                                        pivotCache,
                                        cancellationToken,
                                        $"PivotTable '{pivotName}'");
                                    cancellationToken.ThrowIfCancellationRequested();

                                    cacheOutcomes.Add(cacheIndex, null);
                                    failure = null;
                                }
                                catch (Exception ex) when (IsRecordableQueryFailure(ex, cancellationToken))
                                {
                                    failure = CreatePivotRefreshFailure(pivotName, ex);
                                    cacheOutcomes.Add(cacheIndex, failure);
                                }
                            }

                            if (failure == null)
                            {
                                result.RefreshedPivotTables.Add(pivotName);
                            }
                            else
                            {
                                result.FailedPivotTables.Add(new PowerQueryPivotRefreshFailure
                                {
                                    PivotTableName = pivotName,
                                    ErrorMessage = failure.ErrorMessage,
                                    ExceptionType = failure.ExceptionType,
                                    HResult = failure.HResult
                                });
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref pivotCache);
                            ComUtilities.Release(ref pivotTable);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref pivotTables);
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }
    }

    private static (HashSet<string> TableNames, Dictionary<string, List<WorksheetRangeBounds>> TableRanges)
        ReadWorksheetTableSourcesForQueries(
        Excel.Workbook workbook,
        HashSet<string> queryNames,
        CancellationToken cancellationToken)
    {
        var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tableRanges = new Dictionary<string, List<WorksheetRangeBounds>>(StringComparer.OrdinalIgnoreCase);
        Excel.Sheets? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            for (int i = 1; i <= worksheets.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.Worksheet? worksheet = null;
                Excel.ListObjects? tables = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.Item[i];
                    tables = worksheet.ListObjects;
                    for (int j = 1; j <= tables.Count; j++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        Excel.ListObject? table = null;
                        Excel.QueryTable? queryTable = null;
                        Excel.Range? tableRange = null;
                        try
                        {
                            table = tables.Item[j];
                            if (table.SourceType is not (Excel.XlListObjectSourceType.xlSrcExternal or Excel.XlListObjectSourceType.xlSrcQuery))
                                continue;
                            queryTable = table.QueryTable;

                            object? connectionValue = queryTable.Connection;
                            string? connectionString = connectionValue?.ToString();
                            if (queryNames.Any(queryName =>
                                PowerQuery.PowerQueryHelpers.MatchesMashupLocation(
                                    connectionString,
                                    queryName)))
                            {
                                tableNames.Add(table.Name);
                                tableRange = table.Range;
                                if (!tableRanges.TryGetValue(worksheet.Name, out var ranges))
                                {
                                    ranges = [];
                                    tableRanges.Add(worksheet.Name, ranges);
                                }

                                ranges.Add(ReadRangeBounds(tableRange));
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref tableRange);
                            ComUtilities.Release(ref queryTable);
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
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }

        return (tableNames, tableRanges);
    }

    private static bool HasDataModelLoadForQueries(
        Excel.Workbook workbook,
        HashSet<string> queryNames,
        CancellationToken cancellationToken)
    {
        Excel.Model? model = null;
        Excel.ModelTables? modelTables = null;
        try
        {
            model = workbook.Model;
            modelTables = model.ModelTables;

            for (int i = 1; i <= modelTables.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.ModelTable? modelTable = null;
                Excel.WorkbookConnection? sourceConnection = null;
                try
                {
                    modelTable = modelTables.Item(i);
                    string tableName = modelTable.Name;
                    sourceConnection = modelTable.SourceWorkbookConnection;
                    if (queryNames.Contains(tableName) &&
                        sourceConnection != null &&
                        MatchesDataModelSourceConnection(sourceConnection, tableName))
                    {
                        return true;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref sourceConnection);
                    ComUtilities.Release(ref modelTable);
                }
            }

            return false;
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }
    }

    private static bool IsPivotCacheBackedByQueryTable(
        Excel.Workbook workbook,
        Excel.PivotCache pivotCache,
        HashSet<string> queryTableNames,
        Dictionary<string, List<WorksheetRangeBounds>> queryTableRanges)
    {
        object? sourceData = pivotCache.SourceData;
        if (sourceData is not string source)
        {
            return false;
        }

        source = source.Trim();
        if (queryTableNames.Contains(source) ||
            queryTableNames.Any(tableName =>
                source.EndsWith($"!{tableName}", StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        if (!TryParseWorksheetRangeSource(source, out string sheetName, out string rangeAddress) ||
            !queryTableRanges.TryGetValue(sheetName, out var loadedRanges))
        {
            return false;
        }

        Excel.Sheets? worksheets = null;
        Excel.Application? application = null;
        Excel.Worksheet? worksheet = null;
        Excel.Range? sourceRange = null;
        try
        {
            worksheets = workbook.Worksheets;
            worksheet = (Excel.Worksheet)worksheets.Item[sheetName];
            application = workbook.Application;
            string sourceAddress = rangeAddress;
            if (IsR1C1RangeAddress(rangeAddress))
            {
                object? convertedFormula = application.ConvertFormula(
                    $"={rangeAddress}",
                    Excel.XlReferenceStyle.xlR1C1,
                    Excel.XlReferenceStyle.xlA1);
                if (convertedFormula is not string a1Reference)
                {
                    throw new InvalidOperationException(
                        $"Excel could not convert PivotTable source range '{rangeAddress}' to A1 notation.");
                }

                sourceAddress = a1Reference.TrimStart('=');
            }

            sourceRange = worksheet.Range[sourceAddress];
            var sourceBounds = ReadRangeBounds(sourceRange);
            return loadedRanges.Any(sourceBounds.Intersects);
        }
        finally
        {
            ComUtilities.Release(ref sourceRange);
            ComUtilities.Release(ref worksheet);
            ComUtilities.Release(ref application);
            ComUtilities.Release(ref worksheets);
        }
    }

    private static bool IsR1C1RangeAddress(string rangeAddress)
    {
        string firstCell = rangeAddress.Split(':', 2)[0].Trim();
        if (firstCell.Length < 3 || char.ToUpperInvariant(firstCell[0]) != 'R')
        {
            return false;
        }

        int columnMarker = firstCell.IndexOf('C', StringComparison.OrdinalIgnoreCase);
        return columnMarker > 1 &&
            (char.IsDigit(firstCell[1]) || firstCell[1] == '[');
    }

    private static WorksheetRangeBounds ReadRangeBounds(Excel.Range range)
    {
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            rows = range.Rows;
            columns = range.Columns;
            int firstRow = range.Row;
            int firstColumn = range.Column;
            return new WorksheetRangeBounds(
                firstRow,
                firstColumn,
                firstRow + rows.Count - 1,
                firstColumn + columns.Count - 1);
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    private readonly record struct WorksheetRangeBounds(
        int FirstRow,
        int FirstColumn,
        int LastRow,
        int LastColumn)
    {
        internal bool Intersects(WorksheetRangeBounds other) =>
            FirstRow <= other.LastRow &&
            LastRow >= other.FirstRow &&
            FirstColumn <= other.LastColumn &&
            LastColumn >= other.FirstColumn;
    }

    private static bool TryParseWorksheetRangeSource(
        string source,
        out string sheetName,
        out string rangeAddress)
    {
        bool inQuotedName = false;
        int separator = -1;
        for (int i = 0; i < source.Length; i++)
        {
            if (source[i] == '\'')
            {
                if (inQuotedName && i + 1 < source.Length && source[i + 1] == '\'')
                {
                    i++;
                    continue;
                }

                inQuotedName = !inQuotedName;
            }
            else if (source[i] == '!' && !inQuotedName)
            {
                separator = i;
                break;
            }
        }

        if (separator <= 0 || separator == source.Length - 1 || inQuotedName)
        {
            sheetName = string.Empty;
            rangeAddress = string.Empty;
            return false;
        }

        string sourceSheetName = source[..separator].Trim();
        if (sourceSheetName.Length >= 2 &&
            sourceSheetName[0] == '\'' &&
            sourceSheetName[^1] == '\'')
        {
            sourceSheetName = sourceSheetName[1..^1].Replace("''", "'", StringComparison.Ordinal);
        }

        sheetName = sourceSheetName;
        rangeAddress = source[(separator + 1)..].Trim();
        return sheetName.Length > 0 && rangeAddress.Length > 0;
    }

    private static bool IsDataModelPivotCache(Excel.PivotCache pivotCache)
    {
        if (pivotCache.SourceType != Excel.XlPivotTableSourceType.xlExternal)
        {
            return false;
        }

        Excel.WorkbookConnection? connection = null;
        try
        {
            connection = pivotCache.WorkbookConnection;
            return connection?.Type == Excel.XlConnectionType.xlConnectionTypeMODEL;
        }
        finally
        {
            ComUtilities.Release(ref connection);
        }
    }

    private static PowerQueryPivotRefreshFailure CreatePivotRefreshFailure(
        string pivotName,
        Exception exception) =>
        new()
        {
            PivotTableName = pivotName,
            ErrorMessage = exception.Message,
            ExceptionType = exception.GetType().Name,
            HResult = $"0x{exception.HResult:X8}"
        };

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
