namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal enum WorkbookRefreshState
{
    Ready,
    Refreshing,
    Busy,
    Unknown,
    DialogOpen
}

internal interface IExcelBatchRefreshState
{
    WorkbookRefreshState GetRefreshState();
}

internal sealed class ExcelBusyException(string message, Exception? innerException = null)
    : InvalidOperationException(message, innerException)
{
    internal const string DialogGuidance =
        "Excel is waiting for input in a dialog. Check the Excel window and respond there. " +
        "Ask the user to respond; readable dialog text is included in session status when available. " +
        "ExcelMcp does not identify the dialog type, answer, or dismiss it.";

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
            WorkbookRefreshState.DialogOpen => DialogGuidance,
            _ => "Excel's refresh state could not be confirmed."
        };
        throw new ExcelBusyException(
            $"Cannot {operation}: {reason} Changes have not been saved or discarded; the workbook remains open. " +
            "Check the session again after the dialog or operation finishes, then retry.");
    }
}
