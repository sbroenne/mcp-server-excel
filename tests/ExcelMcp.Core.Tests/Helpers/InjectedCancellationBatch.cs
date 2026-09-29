using Microsoft.Extensions.Logging;
using System.Runtime.ExceptionServices;
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
        Exception? operationError = null;
        inner.Execute((context, _) =>
        {
            try
            {
                operation(context, injectedToken);
            }
            catch (Exception ex)
            {
                operationError = ex;
            }
        });

        if (operationError is not null)
        {
            ExceptionDispatchInfo.Capture(operationError).Throw();
        }
    }

    public T Execute<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default)
    {
        T result = default!;
        Exception? operationError = null;
        inner.Execute((context, _) =>
        {
            try
            {
                result = operation(context, injectedToken);
            }
            catch (Exception ex)
            {
                operationError = ex;
            }
        });

        if (operationError is not null)
        {
            ExceptionDispatchInfo.Capture(operationError).Throw();
        }

        return result;
    }

    public Excel.Workbook GetWorkbook(string filePath) => inner.GetWorkbook(filePath);

    public bool IsExcelProcessAlive() => inner.IsExcelProcessAlive();

    public void Save(CancellationToken cancellationToken = default) =>
        inner.Save(cancellationToken);

    public void UpdateWorkbookPath(string workbookPath) =>
        inner.UpdateWorkbookPath(workbookPath);
}
