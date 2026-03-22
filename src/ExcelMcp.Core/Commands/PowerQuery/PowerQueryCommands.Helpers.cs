using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

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
    /// Strategy order ensures we use the appropriate method for each connection type:
    /// 1. Try QueryTable.Refresh() first (handles worksheet queries)
    /// 2. Fall back to Connection.Refresh() (handles Data Model queries)
    /// </summary>
    /// <returns>True if refresh was executed, false if no connection or table found</returns>
    /// <exception cref="Exception">Thrown if Power Query has formula errors</exception>
    private static bool RefreshConnectionByQueryName(dynamic workbook, string queryName, CancellationToken cancellationToken)
    {
        // Strategy 1: Find and refresh QueryTable directly on worksheet
        // For worksheet queries (InModel=false), errors are thrown by QueryTable.Refresh()
        if (RefreshQueryTableByName(workbook, queryName, cancellationToken))
        {
            return true;
        }

        // Strategy 2: Find connection by name patterns and refresh
        // For Data Model queries (InModel=true), errors are thrown by Connection.Refresh()
        dynamic? targetConnection = null;
        dynamic? connections = null;
        try
        {
            connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
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

    /// <summary>
    /// Finds and refreshes a QueryTable by searching ListObjects on all worksheets.
    /// Matches by query name in the QueryTable's connection string (Location=queryName).
    /// </summary>
    /// <returns>True if QueryTable was found and refreshed</returns>
    /// <exception cref="Exception">Thrown if Power Query has formula errors</exception>
    private static bool RefreshQueryTableByName(dynamic workbook, string queryName, CancellationToken cancellationToken)
    {
        dynamic? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;

            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic? worksheet = null;
                dynamic? listObjects = null;
                try
                {
                    worksheet = worksheets.Item(ws);
                    listObjects = worksheet.ListObjects;

                    for (int lo = 1; lo <= listObjects.Count; lo++)
                    {
                        dynamic? listObject = null;
                        dynamic? queryTable = null;
                        try
                        {
                            listObject = listObjects.Item(lo);

                            // Try to get QueryTable - not all ListObjects have one
                            try
                            {
                                queryTable = listObject.QueryTable;
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                                // ListObject doesn't have a QueryTable - expected for user-created tables
                                continue;
                            }

                            if (queryTable == null)
                            {
                                continue;
                            }

                            // Check if this QueryTable is for our query by examining connection string
                            // Format: "OLEDB;...;Location=QueryName;..."
                            string? connection = queryTable.Connection?.ToString();
                            if (connection != null &&
                                connection.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase))
                            {
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
                                OleMessageFilter.SetPendingCancellationToken(cancellationToken);
                                try
                                {
                                    queryTable.Refresh(false);
                                    SessionDiagnostics.WriteStdErr($"[DIAG-PQ-QT-REFRESH-EXIT-SUCCESS] Query='{queryName}'");
                                }
                                catch (Exception ex)
                                {
                                    SessionDiagnostics.WriteStdErr($"[DIAG-PQ-QT-REFRESH-EXIT-EXCEPTION] Query='{queryName}' ExceptionType={ex.GetType().Name} Message={ex.Message}");
                                    throw;
                                }
                                finally
                                {
                                    OleMessageFilter.ClearPendingCancellationToken();
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

    private static void RefreshWorkbookConnection(dynamic connection, CancellationToken cancellationToken)
    {
        dynamic? oleDbConnection = null;
        bool originalBackgroundQuery = false;
        bool canRestoreBackgroundQuery = false;
        bool supportsRefreshing = false;

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
            catch (COMException)
            {
                // Not an OLEDB connection or provider doesn't support BackgroundQuery.
            }
            catch (RuntimeBinderException)
            {
                // Sub-connection doesn't expose BackgroundQuery via dynamic binding.
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
            string connName = "";
            try { connName = connection.Name?.ToString() ?? "(unknown)"; } catch { }
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
                _ = connection.Refreshing;
                supportsRefreshing = true;
            }
            catch (RuntimeBinderException)
            {
                supportsRefreshing = false;
            }
            catch (COMException)
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
                            return connection.Refreshing;
                        }
                        catch (RuntimeBinderException)
                        {
                            return false;
                        }
                        catch (COMException)
                        {
                            return false;
                        }
                    },
                    () =>
                    {
                        try
                        {
                            connection.CancelRefresh();
                        }
                        catch (RuntimeBinderException)
                        {
                            // Ignore inability to cancel for unsupported providers.
                        }
                        catch (COMException)
                        {
                            // Ignore inability to cancel for unsupported providers.
                        }
                    },
                    cancellationToken);
            }
        }
        finally
        {
            if (canRestoreBackgroundQuery && oleDbConnection != null)
            {
                try
                {
                    oleDbConnection.BackgroundQuery = originalBackgroundQuery;
                }
                catch (COMException)
                {
                    // Ignore inability to restore provider-specific setting.
                }
            }

            ComUtilities.Release(ref oleDbConnection);
        }
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


