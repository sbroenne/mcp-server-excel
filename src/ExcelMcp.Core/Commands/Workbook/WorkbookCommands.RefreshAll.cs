using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    /// <inheritdoc />
    public WorkbookRefreshAllResult RefreshAll(
        IExcelBatch batch,
        TimeSpan timeout,
        IProgress<ProgressInfo>? progress = null)
    {
        timeout = global::Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands.NormalizeRefreshTimeout(timeout);
        var result = new WorkbookRefreshAllResult { FilePath = batch.WorkbookPath };

        return ExcelBatchWait.Execute(batch, timeout, (ctx, ct) =>
        {
            var queryNames = global::Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands.ReadQueryNames(
                ctx.Book,
                ct);
            var queryResult = global::Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands.RefreshQueries(
                queryNames,
                queryName =>
                    global::Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands.RefreshConnectionByQueryName(
                        ctx.Book,
                        queryName,
                        ct),
                ct,
                progress);
            result.RefreshedQueries.AddRange(queryResult.RefreshedQueries);
            result.SkippedQueries.AddRange(queryResult.SkippedQueries);
            result.FailedQueries.AddRange(queryResult.FailedQueries);

            RefreshNonPowerQueryConnections(ctx.Book, result, ct, progress);
            RefreshAllPivotTables(ctx.Book, result, ct, progress);
            FinalizeWorkbookRefreshAllResult(result);
            return result;
        });
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Reports each refresh failure with its connection name, original message, exception type and HRESULT; cancellation and disconnection are not recordable.")]
    private static void RefreshNonPowerQueryConnections(
        Excel.Workbook workbook,
        WorkbookRefreshAllResult result,
        CancellationToken cancellationToken,
        IProgress<ProgressInfo>? progress)
    {
        Excel.Connections? connections = null;
        try
        {
            connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.WorkbookConnection? connection = null;
                try
                {
                    connection = connections.Item(i);
                    if (PowerQueryHelpers.IsPowerQueryConnection(connection))
                    {
                        continue;
                    }

                    string name = connection.Name;
                    progress?.Report(new ProgressInfo
                    {
                        Current = result.RefreshedConnections.Count + result.FailedConnections.Count,
                        Total = connections.Count,
                        Message = $"Refreshing workbook connection '{name}'"
                    });

                    try
                    {
                        global::Sbroenne.ExcelMcp.Core.Commands.ConnectionCommands
                            .RefreshWorkbookConnection(connection, cancellationToken);
                        result.RefreshedConnections.Add(name);
                    }
                    catch (Exception ex) when (
                        global::Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands
                            .IsRecordableQueryFailure(ex, cancellationToken))
                    {
                        result.FailedConnections.Add(new WorkbookRefreshFailure
                        {
                            ConnectionName = name,
                            ErrorMessage = ex.Message,
                            ExceptionType = ex.GetType().Name,
                            HResult = $"0x{ex.HResult:X8}"
                        });
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connection);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref connections);
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Reports each cache refresh failure on its affected pivots with the original message, exception type and HRESULT; cancellation and disconnection propagate.")]
    private static void RefreshAllPivotTables(
        Excel.Workbook workbook,
        WorkbookRefreshAllResult result,
        CancellationToken cancellationToken,
        IProgress<ProgressInfo>? progress)
    {
        var cacheOutcomes = new Dictionary<int, WorkbookPivotRefreshFailure?>();
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
                            string pivotName = pivotTable.Name;
                            int cacheIndex = pivotCache.Index;
                            if (!cacheOutcomes.TryGetValue(cacheIndex, out var failure))
                            {
                                progress?.Report(new ProgressInfo
                                {
                                    Current = result.RefreshedPivotTables.Count + result.FailedPivotTables.Count,
                                    Total = result.RefreshedPivotTables.Count + result.FailedPivotTables.Count + 1,
                                    Message = $"Refreshing PivotTable '{sheetName}!{pivotName}'"
                                });

                                try
                                {
                                    PivotTableRefreshHelper.RefreshSynchronously(
                                        pivotTable,
                                        pivotCache,
                                        cancellationToken,
                                        $"PivotTable '{sheetName}!{pivotName}'");
                                    cancellationToken.ThrowIfCancellationRequested();

                                    cacheOutcomes.Add(cacheIndex, null);
                                    failure = null;
                                }
                                catch (Exception ex) when (
                                    global::Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands
                                        .IsRecordableQueryFailure(ex, cancellationToken))
                                {
                                    failure = new WorkbookPivotRefreshFailure
                                    {
                                        SheetName = sheetName,
                                        PivotTableName = pivotName,
                                        ErrorMessage = ex.Message,
                                        ExceptionType = ex.GetType().Name,
                                        HResult = $"0x{ex.HResult:X8}"
                                    };
                                    cacheOutcomes.Add(cacheIndex, failure);
                                }
                            }

                            if (failure == null)
                            {
                                result.RefreshedPivotTables.Add($"{sheetName}!{pivotName}");
                            }
                            else
                            {
                                result.FailedPivotTables.Add(new WorkbookPivotRefreshFailure
                                {
                                    SheetName = sheetName,
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

    private static void FinalizeWorkbookRefreshAllResult(WorkbookRefreshAllResult result)
    {
        int failures =
            result.FailedQueries.Count +
            result.FailedConnections.Count +
            result.FailedPivotTables.Count;
        int refreshed =
            result.RefreshedQueries.Count +
            result.RefreshedConnections.Count +
            result.RefreshedPivotTables.Count;
        result.Message =
            $"Workbook refresh-all refreshed {result.RefreshedQueries.Count} Power Query load(s), " +
            $"{result.RefreshedConnections.Count} other connection(s), and " +
            $"{result.RefreshedPivotTables.Count} PivotTable(s); " +
            $"{result.SkippedQueries.Count} Power Query definition(s) had nothing to refresh independently.";

        if (failures == 0)
        {
            result.Success = true;
            result.ErrorMessage = null;
            return;
        }

        result.Success = false;
        result.ErrorMessage =
            $"Workbook refresh-all refreshed {refreshed} object(s) and encountered {failures} failure(s). " +
            "Successful refreshes were not rolled back. Inspect failedQueries, failedConnections, " +
            "and failedPivotTables for details. " + result.Message;
    }
}
