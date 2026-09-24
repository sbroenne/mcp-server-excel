using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

internal static class ConnectionRefreshHelpers
{
    internal static void RefreshWorkbookConnection(
        Excel.WorkbookConnection connection,
        CancellationToken cancellationToken,
        Action<Func<bool>, Action, CancellationToken> waitForCompletion)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Excel.OLEDBConnection? oledb = null;
        Excel.ODBCConnection? odbc = null;
        try
        {
            switch (connection.Type)
            {
                case Excel.XlConnectionType.xlConnectionTypeOLEDB:
                    oledb = connection.OLEDBConnection;
                    Refresh(
                        () => connection.Refresh(), () => oledb.BackgroundQuery, value => oledb.BackgroundQuery = value,
                        () => oledb.Refreshing, () => oledb.CancelRefresh(), cancellationToken, waitForCompletion);
                    break;
                case Excel.XlConnectionType.xlConnectionTypeODBC:
                    odbc = connection.ODBCConnection;
                    Refresh(
                        () => connection.Refresh(), () => odbc.BackgroundQuery, value => odbc.BackgroundQuery = value,
                        () => odbc.Refreshing, () => odbc.CancelRefresh(), cancellationToken, waitForCompletion);
                    break;
                default:
                    // Other types have no OLEDB/ODBC status API. Retain their native Refresh
                    // behavior without probing nonexistent WorkbookConnection members.
                    RefreshWithCancellation(() => connection.Refresh(), cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    break;
            }
        }
        finally
        {
            ComUtilities.Release(ref odbc);
            ComUtilities.Release(ref oledb);
        }
    }

    private static void Refresh(
        Action refresh,
        Func<bool> getBackgroundQuery,
        Action<bool> setBackgroundQuery,
        Func<bool> isRefreshing,
        Action cancelRefresh,
        CancellationToken cancellationToken,
        Action<Func<bool>, Action, CancellationToken> waitForCompletion)
    {
        bool originalBackgroundQuery = false;
        bool backgroundQueryChanged = false;
        bool completed = false;
        try
        {
            try
            {
                originalBackgroundQuery = getBackgroundQuery();
                // OLAP providers expose a read-only false value; avoid unnecessary writes.
                if (originalBackgroundQuery)
                {
                    setBackgroundQuery(false);
                    backgroundQueryChanged = true;
                }
            }
            catch (COMException)
            {
                // Some providers cannot switch modes; typed status polling remains required.
            }

            RefreshWithCancellation(refresh, cancellationToken);
            waitForCompletion(isRefreshing, cancelRefresh, cancellationToken);
            completed = true;
        }
        finally
        {
            if (backgroundQueryChanged)
            {
                try
                {
                    setBackgroundQuery(originalBackgroundQuery);
                }
                catch (COMException) when (!completed)
                {
                    // A secondary restore failure must not hide the refresh/status failure.
                }
            }
        }
    }

    private static void RefreshWithCancellation(Action refresh, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // Synchronous refresh needs inbound Excel callbacks; do not use EnterLongOperation.
        // Restore the enclosing batch operation token so later polling/cancellation calls stay cancellable.
        var previousToken = OleMessageFilter.ExchangePendingCancellationToken(cancellationToken);
        try
        {
            refresh();
        }
        finally
        {
            OleMessageFilter.ExchangePendingCancellationToken(previousToken);
        }
    }

    internal static void EnsureQueryTableRefreshSucceeded(bool refreshed)
    {
        if (!refreshed)
            throw new InvalidOperationException("QueryTable refresh was cancelled before completion.");
    }
}
