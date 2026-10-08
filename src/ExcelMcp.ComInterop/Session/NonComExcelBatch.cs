using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

// Keep the embedded-PIA interface slots in their defining assembly. Generic COM
// signatures cannot be implemented by independently embedded equivalent types.
internal abstract class NonComExcelBatch : IExcelBatch
{
    public abstract string WorkbookPath { get; }
    public abstract TimeSpan OperationTimeout { get; }
    public abstract bool IsExcelVisible { get; }
    public abstract bool HasTimedOutOperation { get; }
    public ILogger Logger => NullLogger.Instance;
    public int? ExcelProcessId => null;
    public IReadOnlyDictionary<string, Excel.Workbook> Workbooks => throw ComUnavailable();

    public Excel.Workbook GetWorkbook(string filePath) => throw ComUnavailable();
    public void Execute(Action<ExcelContext, CancellationToken> operation, CancellationToken cancellationToken = default) => throw ComUnavailable();
    public T Execute<T>(Func<ExcelContext, CancellationToken, T> operation, CancellationToken cancellationToken = default) => throw ComUnavailable();
    public bool IsExcelProcessAlive() => throw new PlatformNotSupportedException("This batch owns a workbook, not the shared Excel process.");
    public abstract void UpdateWorkbookPath(string workbookPath);
    public abstract void Save(CancellationToken cancellationToken = default);
    public abstract void Dispose();

    private static PlatformNotSupportedException ComUnavailable() => new("COM callbacks and workbook references are not available in a non-COM workbook batch.");
}
