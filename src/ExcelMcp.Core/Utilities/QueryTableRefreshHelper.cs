using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class QueryTableRefreshHelper
{
    internal static void RefreshSynchronously(
        Excel.QueryTable queryTable,
        CancellationToken cancellationToken,
        string operation)
    {
        OleMessageFilter.SetPendingCancellationToken(cancellationToken);
        try
        {
            EnsureSucceeded(queryTable.Refresh(false), operation);
        }
        finally
        {
            OleMessageFilter.ClearPendingCancellationToken();
        }
    }

    internal static void EnsureSucceeded(bool refreshed, string operation)
    {
        if (!refreshed)
        {
            throw new OperationFailureException(
                OperationFailureCategory.Cancelled,
                $"{operation} was cancelled by Excel.");
        }
    }
}
