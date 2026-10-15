using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class PivotTableRefreshHelper
{
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Retains refresh and BackgroundQuery restoration failures, then throws the original failure or their aggregate after restoration.")]
    internal static void RefreshSynchronously(
        Excel.PivotTable pivotTable,
        Excel.PivotCache pivotCache,
        CancellationToken cancellationToken,
        string operation)
    {
        bool originalBackgroundQuery = false;
        bool canRestoreBackgroundQuery = false;
        Exception? refreshFailure = null;
        COMException? restoreFailure = null;

        try
        {
            if (pivotCache.SourceType == Excel.XlPivotTableSourceType.xlExternal && !pivotCache.OLAP)
            {
                originalBackgroundQuery = pivotCache.BackgroundQuery;
                canRestoreBackgroundQuery = true;
                pivotCache.BackgroundQuery = false;
            }

            OleMessageFilter.SetPendingCancellationToken(cancellationToken);
            try
            {
                if (!pivotTable.RefreshTable())
                {
                    throw new InvalidOperationException(
                        $"Excel reported that {operation} did not refresh.");
                }
            }
            finally
            {
                OleMessageFilter.ClearPendingCancellationToken();
            }
        }
        catch (Exception ex)
        {
            refreshFailure = ex;
        }
        finally
        {
            if (canRestoreBackgroundQuery)
            {
                try
                {
                    pivotCache.BackgroundQuery = originalBackgroundQuery;
                }
                catch (COMException ex)
                {
                    restoreFailure = ex;
                }
            }
        }

        if (restoreFailure != null)
        {
            throw new InvalidOperationException(
                $"Could not restore the PivotTable cache's original BackgroundQuery setting after {operation}.",
                refreshFailure == null
                    ? restoreFailure
                    : new AggregateException(refreshFailure, restoreFailure));
        }

        if (refreshFailure != null)
        {
            ExceptionDispatchInfo.Capture(refreshFailure).Throw();
        }
    }
}
