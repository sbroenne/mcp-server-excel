using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query helper methods (internal utilities)
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Core connection refresh logic - finds and refreshes the connection for a query.
    ///
    /// Error propagation depends on connection type:
    /// - Worksheet queries (InModel=false): Errors thrown via QueryTable.Refresh(false)
    /// - Data Model queries (InModel=true): Errors thrown via Connection.Refresh()
    ///
    /// Strategy order ensures every configured destination is refreshed:
    /// 1. Refresh the worksheet QueryTable when present
    /// 2. Refresh the Data Model source connection when present
    /// 3. Fall back to a matching workbook connection for connection-only queries
    /// </summary>
    /// <returns>True if refresh was executed, false if no connection or table found</returns>
    /// <exception cref="Exception">Thrown if Power Query has formula errors</exception>
    internal static bool RefreshConnectionByQueryName(Excel.Workbook workbook, string queryName, CancellationToken cancellationToken)
    {
        // A LoadToBoth query has independent worksheet and Data Model connections.
        // Refresh both before reporting success so neither destination remains stale.
        var worksheetRefreshed = RefreshQueryTableByName(
            workbook,
            queryName,
            cancellationToken);
        var dataModelRefreshed = RefreshDataModelConnectionByQueryName(
            workbook,
            queryName,
            cancellationToken);
        if (worksheetRefreshed || dataModelRefreshed)
        {
            return true;
        }

        // Connection-only queries may expose a workbook connection without either
        // a worksheet QueryTable or a Data Model table.
        Excel.WorkbookConnection? targetConnection = null;
        Excel.Connections? connections = null;
        try
        {
            connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.WorkbookConnection? conn = null;
                try
                {
                    conn = connections[i];
                    if (PowerQuery.PowerQueryHelpers.TryGetMashupLocation(
                        conn,
                        out var location) &&
                        string.Equals(location, queryName, StringComparison.OrdinalIgnoreCase))
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
        finally
        {
            ComUtilities.Release(ref connections);
        }

        if (targetConnection != null)
        {
            try
            {
                RefreshWorkbookConnection(targetConnection, cancellationToken);
                return true;
            }
            finally
            {
                ComUtilities.Release(ref targetConnection);
            }
        }

        return false;
    }

    private static bool RefreshDataModelConnectionByQueryName(
        Excel.Workbook workbook,
        string queryName,
        CancellationToken cancellationToken)
    {
        Excel.Model? model = null;
        Excel.ModelTables? modelTables = null;
        try
        {
            try
            {
                model = workbook.Model;
                modelTables = model.ModelTables;
            }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
            {
                return false;
            }

            for (int i = 1; i <= modelTables.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.ModelTable? modelTable = null;
                Excel.WorkbookConnection? sourceConnection = null;
                try
                {
                    modelTable = modelTables[i];
                    string tableName = modelTable.Name?.ToString() ?? string.Empty;
                    sourceConnection = modelTable.SourceWorkbookConnection;
                    if (string.Equals(
                            tableName,
                            queryName,
                            StringComparison.OrdinalIgnoreCase) &&
                        sourceConnection != null &&
                        MatchesDataModelSourceConnection(sourceConnection, queryName))
                    {
                        RefreshWorkbookConnection(sourceConnection, cancellationToken);
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

    /// <summary>
    /// Finds and refreshes a QueryTable by searching ListObjects on all worksheets.
    /// Matches by query name in the QueryTable's connection string (Location=queryName).
    /// </summary>
    /// <returns>True if QueryTable was found and refreshed</returns>
    /// <exception cref="Exception">Thrown if Power Query has formula errors</exception>
    private static bool RefreshQueryTableByName(Excel.Workbook workbook, string queryName, CancellationToken cancellationToken)
    {
        Excel.Sheets? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;

            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                Excel.Worksheet? worksheet = null;
                Excel.ListObjects? listObjects = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets[ws];
                    listObjects = worksheet.ListObjects;

                    for (int lo = 1; lo <= listObjects.Count; lo++)
                    {
                        Excel.ListObject? listObject = null;
                        Excel.QueryTable? queryTable = null;
                        try
                        {
                            listObject = listObjects[lo];

                            var sourceType = listObject.SourceType;
                            if (sourceType is not (Excel.XlListObjectSourceType.xlSrcExternal or Excel.XlListObjectSourceType.xlSrcQuery))
                                continue;
                            queryTable = listObject.QueryTable;

                            if (queryTable == null)
                            {
                                continue;
                            }

                            // Check if this QueryTable is for our query by examining connection string
                            // Format: "OLEDB;...;Location=QueryName;..."
                            object? connectionValue = queryTable.Connection;
                            string? connection = connectionValue?.ToString();
                            if (PowerQuery.PowerQueryHelpers.MatchesMashupLocation(
                                connection,
                                queryName))
                            {
                                // Also repair bindings in tables loaded by older server versions.
                                queryTable.PreserveColumnInfo = true;
                                // Keep synchronous refresh semantics for worksheet queries.
                                // QueryTable.Refresh(false) is the only reliable path that propagates
                                // Power Query formula errors for worksheet-loaded queries.
                                //
                                // Do NOT use EnterLongOperation() here. EnterLongOperation sets
                                // _isInLongOperation=true, causing HandleInComingCall to return
                                // SERVERCALL_RETRYLATER for all inbound COM calls — including any
                                // callbacks Excel needs to complete the synchronous refresh.
                                // This can create a mutual deadlock. Instead, register the
                                // CancellationToken so MessagePending returns PENDINGMSG_CANCELCALL
                                // if cancelled, allowing the STA thread to exit cleanly.

                                // INSTRUMENTATION: Trace QueryTable refresh entry and exit
                                SessionDiagnostics.WriteStdErr($"[DIAG-PQ-QT-REFRESH-ENTER] Query='{queryName}' WorksheetIndex={ws}");
                                try
                                {
                                    QueryTableRefreshHelper.RefreshSynchronously(
                                        queryTable,
                                        cancellationToken,
                                        $"Power Query refresh for '{queryName}'");
                                    SessionDiagnostics.WriteStdErr($"[DIAG-PQ-QT-REFRESH-EXIT-SUCCESS] Query='{queryName}'");
                                }
                                catch (Exception ex)
                                {
                                    SessionDiagnostics.WriteStdErr($"[DIAG-PQ-QT-REFRESH-EXIT-EXCEPTION] Query='{queryName}' ExceptionType={ex.GetType().Name} Message={ex.Message}");
                                    throw;
                                }
                                return true;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref queryTable);
                            ComUtilities.Release(ref listObject);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listObjects);
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }

        return false;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Collects refresh and BackgroundQuery restoration failures, releases COM references, then rethrows the primary error or both failures together.")]
    private static void RefreshWorkbookConnection(Excel.WorkbookConnection connection, CancellationToken cancellationToken)
    {
        Excel.OLEDBConnection? oleDbConnection = null;
        bool originalBackgroundQuery = false;
        bool canRestoreBackgroundQuery = false;
        bool supportsRefreshing = false;
        Exception? refreshFailure = null;
        COMException? restoreFailure = null;

        try
        {
            try
            {
                oleDbConnection = connection.OLEDBConnection;
                if (oleDbConnection != null)
                {
                    originalBackgroundQuery = oleDbConnection.BackgroundQuery;
                    canRestoreBackgroundQuery = true;

                    // CRITICAL: Force BackgroundQuery = false to ensure synchronous refresh.
                    //
                    // With BackgroundQuery = true (async), connection.Refresh() returns immediately
                    // while Excel processes the query in a background thread. We then poll
                    // connection.Refreshing with Thread.Sleep(200). On STA threads with the
                    // OleMessageFilter registered, COM events from Excel during background refresh
                    // (SheetChange, Calculate, Data Model callbacks) cause Thread.Sleep to return
                    // via MsgWaitForMultipleObjectsEx — turning the polling loop into a 100% CPU
                    // spin lasting the full duration of the refresh (seconds to minutes).
                    //
                    // With BackgroundQuery = false (synchronous), connection.Refresh() blocks the
                    // STA thread until the refresh completes. When it returns, connection.Refreshing
                    // is already false, so WaitForRefreshCompletion exits in 0 iterations. Zero spin.
                    oleDbConnection.BackgroundQuery = false;
                }
            }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
            {
                // Not an OLEDB connection or provider doesn't support BackgroundQuery.
            }

            // IMPORTANT: Do NOT use EnterLongOperation() here.
            //
            // connection.Refresh() with BackgroundQuery=false is a synchronous COM call that
            // requires Excel to callback into our STA apartment to complete the data load.
            // EnterLongOperation() sets _isInLongOperation=true, which causes HandleInComingCall
            // to return SERVERCALL_RETRYLATER for ALL inbound COM calls — including the essential
            // callbacks Excel needs to complete connection.Refresh(). This creates a mutual
            // deadlock: Excel waits for our callbacks to be accepted; our STA thread waits for
            // Excel to respond. The thread hangs indefinitely (observed: 30-minute hang).
            //
            // Instead, register the CancellationToken with the message filter so MessagePending
            // returns PENDINGMSG_CANCELCALL (0) when cancelled, causing connection.Refresh() to
            // return RPC_E_CALL_CANCELLED and unblocking the STA thread cleanly.
            //
            // Trade-off: without EnterLongOperation, inbound EnsureScanDefinedEvents callbacks
            // are not throttled, which may cause elevated CPU during refresh (~88% peak).
            // This is preferable to a permanent hang.

            // INSTRUMENTATION: Trace Connection refresh entry and exit
            string connName = connection.Name?.ToString() ?? "(unknown)";
            SessionDiagnostics.WriteStdErr($"[DIAG-PQ-CONN-REFRESH-ENTER] Connection='{connName}'");
            OleMessageFilter.SetPendingCancellationToken(cancellationToken);
            try
            {
                connection.Refresh();
                SessionDiagnostics.WriteStdErr($"[DIAG-PQ-CONN-REFRESH-EXIT-SUCCESS] Connection='{connName}'");
            }
            catch (Exception ex)
            {
                SessionDiagnostics.WriteStdErr($"[DIAG-PQ-CONN-REFRESH-EXIT-EXCEPTION] Connection='{connName}' ExceptionType={ex.GetType().Name} Message={ex.Message}");
                throw;
            }
            finally
            {
                OleMessageFilter.ClearPendingCancellationToken();
            }

            try
            {
                supportsRefreshing = oleDbConnection is not null;
                if (supportsRefreshing && oleDbConnection is not null)
                {
                    _ = oleDbConnection.Refreshing;
                }
            }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
            {
                supportsRefreshing = false;
            }

            if (supportsRefreshing)
            {
                WaitForRefreshCompletion(
                    () =>
                    {
                        try
                        {
                            return oleDbConnection!.Refreshing;
                        }
                        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
                        {
                            return false;
                        }
                    },
                    () =>
                    {
                        try
                        {
                            oleDbConnection!.CancelRefresh();
                        }
                        catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
                        {
                            // Ignore inability to cancel for unsupported providers.
                        }
                    },
                    cancellationToken);
            }
        }
        catch (Exception exception)
        {
            refreshFailure = exception;
        }
        finally
        {
            try
            {
                if (canRestoreBackgroundQuery && oleDbConnection is not null)
                    oleDbConnection.BackgroundQuery = originalBackgroundQuery;
            }
            catch (COMException restorationError)
            {
                restoreFailure = restorationError;
            }
            finally
            {
                ComUtilities.Release(ref oleDbConnection);
            }
        }
        if (restoreFailure is not null)
        {
            var failure = new AggregateException("Power Query refresh could not restore BackgroundQuery.",
                refreshFailure is null ? [restoreFailure] : [refreshFailure, restoreFailure]);
            if (refreshFailure is OperationCanceledException canceled)
                throw new OperationCanceledException(failure.Message, failure, canceled.CancellationToken);
            throw failure;
        }
        if (refreshFailure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(refreshFailure).Throw();
    }

    private static void WaitForRefreshCompletion(
        Func<bool> isRefreshing,
        Action cancelRefresh,
        CancellationToken cancellationToken)
    {
        // CRITICAL: Rate-limit the isRefreshing() COM call to every 200ms of *real* elapsed time.
        //
        // On STA threads with OleMessageFilter registered, COM events from Excel during refresh
        // (SheetChange, Calculate, Data Model callbacks) wake Thread.Sleep immediately via
        // MsgWaitForMultipleObjectsEx (CoWaitForMultipleHandles). Without rate-limiting,
        // isRefreshing() (a cross-process COM property access, ~200-500μs) runs thousands of
        // times/second → 100% CPU spin.
        //
        // Use KernelSleep (Win32 Sleep via P/Invoke) instead of Thread.Sleep:
        // Thread.Sleep on STA threads uses CoWaitForMultipleHandles which pumps the COM message
        // queue and wakes early on every incoming COM event (Data Model row-write callbacks from
        // MashupHost.exe, SheetChange, etc.). During large PQ refreshes this causes CPU spin
        // even with the Stopwatch guard. Win32 Sleep() is a bare NtDelayExecution call with no
        // COM pumping — the thread genuinely sleeps the full 200ms per interval.
        // Safety: refresh completion is driven by Excel's own internals (MashupHost → Excel STA).
        // connection.Refreshing flips to false in Excel's process without requiring our STA to
        // service any callbacks. The Stopwatch guard is kept as defensive belt-and-suspenders.
        const int CheckIntervalMs = 200;
        var sw = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            // Initial check: if already done, skip the wait entirely.
            if (!isRefreshing())
                return;

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Sleep without pumping the STA COM queue. Win32 Sleep() does not wake early
                // on COM events, so the Stopwatch guard below is belt-and-suspenders only.
                ComUtilities.KernelSleep(CheckIntervalMs);
                // Guard: loop back without calling isRefreshing() if sleep returned early.
                if (sw.Elapsed.TotalMilliseconds < CheckIntervalMs)
                    continue;
                sw.Restart();
                if (!isRefreshing())
                    break;
            }
        }
        catch (OperationCanceledException)
        {
            cancelRefresh();
            throw;
        }
    }
}
