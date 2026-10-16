namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal enum WorkbookRefreshState
{
    Ready,
    Refreshing,
    Busy,
    Unknown,
    DialogOpen,
    NativeQueryApproval
}

internal interface IExcelBatchRefreshState
{
    WorkbookRefreshState GetRefreshState();
}

internal sealed class ExcelBusyException(string message, Exception? innerException = null)
    : InvalidOperationException(message, innerException)
{
    internal const string NativeQueryApprovalGuidance =
        "Excel is waiting for approval to run a native database query. Show the Excel window and approve the query manually; the operation is still pending.";

    internal static void ThrowIfNotReady(WorkbookRefreshState state, string operation)
    {
        if (state == WorkbookRefreshState.Ready)
        {
            return;
        }

        var reason = state switch
        {
            WorkbookRefreshState.Refreshing => "An Excel data refresh is still running.",
            WorkbookRefreshState.Busy => "Excel is busy. A refresh or another Excel operation may still be running.",
            WorkbookRefreshState.DialogOpen => "Excel has a dialog open. Check the Excel window for a prompt; it may require sign-in or other user input.",
            WorkbookRefreshState.NativeQueryApproval => NativeQueryApprovalGuidance,
            _ => "Excel's refresh state could not be confirmed."
        };
        throw new ExcelBusyException(
            $"Cannot {operation}: {reason} Changes have not been saved or discarded; the workbook remains open. " +
            "Check the session again after the dialog or operation finishes, then retry.");
    }
}
