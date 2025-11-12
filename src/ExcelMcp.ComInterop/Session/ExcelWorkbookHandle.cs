namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Wraps Excel COM objects for a single workbook.
/// Tracks last access time for automatic cleanup.
/// </summary>
public sealed class ExcelWorkbookHandle : IAsyncDisposable, IDisposable
{
    private dynamic? _application;
    private dynamic? _workbook;
    private bool _disposed;

    /// <summary>
    /// Absolute file path of the workbook
    /// </summary>
    public string FilePath { get; }

    /// <summary>
    /// Last time the handle was accessed (for inactivity timeout)
    /// </summary>
    public DateTime LastAccess { get; private set; }

    /// <summary>
    /// Excel Application COM object
    /// </summary>
    public dynamic Application => _application ?? throw new ObjectDisposedException(nameof(ExcelWorkbookHandle));

    /// <summary>
    /// Excel Workbook COM object
    /// </summary>
    public dynamic Workbook => _workbook ?? throw new ObjectDisposedException(nameof(ExcelWorkbookHandle));

    private ExcelWorkbookHandle(string filePath)
    {
        FilePath = Path.GetFullPath(filePath);
        LastAccess = DateTime.UtcNow;
    }

    /// <summary>
    /// Creates a new handle by opening Excel and the workbook
    /// </summary>
    public static async Task<ExcelWorkbookHandle> CreateAsync(string filePath)
    {
        var handle = new ExcelWorkbookHandle(filePath);
        await Task.Run(() =>
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Excel is not installed on this system");

            handle._application = Activator.CreateInstance(excelType);
            handle._application.Visible = false;
            handle._application.DisplayAlerts = false;

            if (File.Exists(filePath))
            {
                handle._workbook = handle._application.Workbooks.Open(filePath);
            }
            else
            {
                handle._workbook = handle._application.Workbooks.Add();
                handle._workbook.SaveAs(filePath);
            }
        });
        return handle;
    }

    /// <summary>
    /// Updates the last access time (resets inactivity timeout)
    /// </summary>
    public void UpdateLastAccess()
    {
        LastAccess = DateTime.UtcNow;
    }

    /// <summary>
    /// Saves the workbook
    /// </summary>
    public async Task SaveAsync()
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(ExcelWorkbookHandle));

        await Task.Run(() => _workbook?.Save());
        UpdateLastAccess();
    }

    /// <summary>
    /// Executes an operation with access to Application and Workbook COM objects.
    /// Provides batch-like context for COM operations.
    /// </summary>
    /// <typeparam name="T">Return type</typeparam>
    /// <param name="operation">Operation to execute with ExcelContext</param>
    /// <param name="timeout">Optional timeout for the operation</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Result of the operation</returns>
    public async Task<T> ExecuteAsync<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(ExcelWorkbookHandle));

        UpdateLastAccess();

        return await Task.Run(() =>
        {
            using var cts = timeout.HasValue
                ? CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                : null;

            if (timeout.HasValue)
            {
                cts!.CancelAfter(timeout.Value);
            }

            var effectiveToken = cts?.Token ?? cancellationToken;
            var context = new ExcelContext(FilePath, _application!, _workbook!);

            return operation(context, effectiveToken);
        }, cancellationToken);
    }

    /// <summary>
    /// Executes an operation with access to Application and Workbook COM objects (non-generic version).
    /// Provides batch-like context for COM operations that return ValueTask.
    /// </summary>
    public ValueTask<T> Execute<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        return new ValueTask<T>(ExecuteAsync(operation, timeout, cancellationToken));
    }

    /// <summary>
    /// Closes the workbook and quits Excel
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        _disposed = true;

        await Task.Run(() =>
        {
            try
            {
                // Close workbook without saving
                if (_workbook != null)
                {
                    try
                    {
                        _workbook.Close(false);
                    }
                    catch { /* Ignore close errors */ }

                    ComUtilities.Release(ref _workbook);
                }

                // Quit Excel application
                if (_application != null)
                {
                    try
                    {
                        _application.Quit();
                    }
                    catch { /* Ignore quit errors */ }

                    ComUtilities.Release(ref _application);
                }

                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch
            {
                // Suppress all disposal errors
            }
        });
    }

    /// <summary>
    /// Disposes the handle synchronously
    /// </summary>
    public void Dispose()
    {
        DisposeAsync().AsTask().GetAwaiter().GetResult();
    }
}
