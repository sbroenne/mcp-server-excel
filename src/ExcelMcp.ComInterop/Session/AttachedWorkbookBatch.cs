using Microsoft.Extensions.Logging;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed class AttachedWorkbookBatch(ExcelBatch owner, OpenWorkbookInfo workbook)
    : IExcelBatch, IExcelBatchRefreshState, IExcelBatchCloseState, IExcelBatchWindowControl
{
    private string _path = workbook.Location ?? workbook.Name;
    private string _workbookName = workbook.Name;
    private bool _detached;
    private long _locationRevision;

    internal ExcelBatch Owner => owner;
    internal string WorkbookId => workbook.WorkbookId;
    internal bool? AutoSaveOn => workbook.AutoSaveOn;
    internal string WorkbookName => _workbookName;
    internal object LocationSyncRoot { get; } = new();
    internal long LocationRevision => Interlocked.Read(ref _locationRevision);
    internal void InvalidateLocationSnapshot()
    {
        lock (LocationSyncRoot)
            Interlocked.Increment(ref _locationRevision);
    }
    internal string? FilePath => WorkbookLocation.IsRemote(_path) || Path.IsPathFullyQualified(_path)
        ? _path
        : null;
    public string WorkbookPath => _path;
    public ILogger Logger => owner.Logger;
    public int? ExcelProcessId => owner.ExcelProcessId;
    public TimeSpan OperationTimeout => owner.OperationTimeout;
    public bool IsExcelVisible => owner.IsExcelVisible;
    public bool HasTimedOutOperation => owner.HasTimedOutOperation;
    public bool ShutdownCommitted => _detached;
    public string? DialogDetails => owner.DialogDetails;
    public string? DialogVisibilityError => owner.DialogVisibilityError;
    public IReadOnlyDictionary<string, Excel.Workbook> Workbooks =>
        new Dictionary<string, Excel.Workbook>(StringComparer.OrdinalIgnoreCase)
        {
            [_path] = GetSelectedWorkbook()
        };

    public Excel.Workbook GetWorkbook(string filePath) =>
        string.Equals(filePath, _path, StringComparison.OrdinalIgnoreCase)
            ? GetSelectedWorkbook()
            : throw new KeyNotFoundException("The requested workbook is not the selected attached workbook.");

    public void UpdateWorkbookPath(string workbookPath) =>
        UpdateDiscoveredWorkbookLocation(workbookPath, Path.GetFileName(workbookPath));

    internal void UpdateDiscoveredWorkbookLocation(string? location, string name)
    {
        var path = location is null
            ? name
            : WorkbookLocation.Normalize(location);
        lock (LocationSyncRoot)
        {
            if (!string.Equals(_path, path, StringComparison.Ordinal)
                || !string.Equals(_workbookName, name, StringComparison.Ordinal))
            {
                InvalidateLocationSnapshot();
            }
            _path = path;
            _workbookName = name;
        }
    }

    public bool IsExcelProcessAlive() => !_detached && owner.IsExcelProcessAlive();
    public ExcelVisibilityResult SetVisibility(bool visible) => owner.SetVisibility(visible);
    public WorkbookRefreshState GetRefreshState() =>
        owner.GetAttachedWorkbookRefreshState(workbook.WorkbookId);

    public T Execute<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_detached, this);
        return owner.Execute((_, token) =>
        {
            var context = owner.GetAttachedWorkbookContext(workbook.WorkbookId);
            var name = context.Book.Name;
            UpdateDiscoveredWorkbookLocation(
                string.Equals(context.WorkbookPath, name, StringComparison.Ordinal) ? null : context.WorkbookPath,
                name);
            return operation(context, token);
        }, cancellationToken);
    }

    public void Execute(
        Action<ExcelContext, CancellationToken> operation,
        CancellationToken cancellationToken = default) =>
        Execute((context, token) =>
        {
            operation(context, token);
            return 0;
        }, cancellationToken);

    public void Save(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ObjectDisposedException.ThrowIf(_detached, this);
        ExcelBusyException.ThrowIfNotReady(owner.GetExecutionPreflightState(), "save", this);
        Execute((context, token) =>
        {
            ExcelBusyException.ThrowIfNotReady(
                WorkbookRefreshProbe.Read(context.App, context.Book), "save");
            ExcelShutdownService.SaveWorkbookWithTimeout(
                context.Book, context.Book.Name, Logger, token);
        }, cancellationToken);
    }

    public void Close()
        => Close(force: false);

    internal void Close(bool force)
    {
        ObjectDisposedException.ThrowIf(_detached, this);
        owner.Execute((_, _) => owner.CloseAttachedWorkbook(workbook.WorkbookId, force));
        _detached = true;
    }

    public void Dispose() => _detached = true;

    private Excel.Workbook GetSelectedWorkbook()
    {
        ObjectDisposedException.ThrowIf(_detached, this);
        return owner.GetAttachedWorkbook(workbook.WorkbookId);
    }
}
