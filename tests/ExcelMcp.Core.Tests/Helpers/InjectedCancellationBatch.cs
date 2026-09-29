using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

internal sealed class InjectedCancellationBatch(
    IExcelBatch inner,
    CancellationToken injectedToken) : IExcelBatch
{
    public string WorkbookPath => inner.WorkbookPath;

    public ILogger Logger => inner.Logger;

    public IReadOnlyDictionary<string, Excel.Workbook> Workbooks => inner.Workbooks;

    public bool HasTimedOutOperation => inner.HasTimedOutOperation;

    public int? ExcelProcessId => inner.ExcelProcessId;

    public TimeSpan OperationTimeout => inner.OperationTimeout;

    public bool IsExcelVisible => inner.IsExcelVisible;

    public void Dispose()
    {
    }

    public void Execute(
        Action<ExcelContext, CancellationToken> operation,
        CancellationToken cancellationToken = default)
    {
        try
        {
            inner.Execute((context, _) => operation(context, injectedToken));
        }
        catch (TimeoutException) when (injectedToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(injectedToken);
        }
    }

    public T Execute<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default)
    {
        try
        {
            return inner.Execute((context, _) => operation(context, injectedToken));
        }
        catch (TimeoutException) when (injectedToken.IsCancellationRequested)
        {
            throw new OperationCanceledException(injectedToken);
        }
    }

    public Excel.Workbook GetWorkbook(string filePath) => inner.GetWorkbook(filePath);

    public bool IsExcelProcessAlive() => inner.IsExcelProcessAlive();

    public void Save(CancellationToken cancellationToken = default) =>
        inner.Save(cancellationToken);

    public void UpdateWorkbookPath(string workbookPath) =>
        inner.UpdateWorkbookPath(workbookPath);
}
