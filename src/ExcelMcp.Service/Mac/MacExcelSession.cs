namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelSession
{
    private readonly object _stateLock = new();
    private TaskCompletionSource? _drained;
    private bool _closing;
    private int _pendingOperations;

    public required string SessionId { get; init; }
    public required string FilePath { get; init; }
    public required TimeSpan OperationTimeout { get; init; }
    public required bool IsVisible { get; set; }
    public bool RequiresRecovery { get; set; }
    public bool HasUnconfirmedOpen { get; set; }
    public string? UnsafeReason { get; private set; }
    public DateTime CreatedAt { get; } = DateTime.UtcNow;
    public SemaphoreSlim OperationLock { get; } = new(1, 1);
    public int ActiveOperations;
    public int PendingOperations
    {
        get
        {
            lock (_stateLock)
            {
                return _pendingOperations;
            }
        }
    }

    public bool IsClosing
    {
        get
        {
            lock (_stateLock)
            {
                return _closing;
            }
        }
    }

    public bool TryAdmitOperation()
    {
        lock (_stateLock)
        {
            if (_closing)
            {
                return false;
            }

            _pendingOperations++;
            return true;
        }
    }

    public void MarkUnsafe(string reason)
    {
        lock (_stateLock)
        {
            UnsafeReason ??= reason;
        }
    }

    public void CompleteOperation()
    {
        TaskCompletionSource? drained = null;
        lock (_stateLock)
        {
            _pendingOperations--;
            if (_closing && _pendingOperations == 0)
            {
                drained = _drained;
            }
        }

        drained?.TrySetResult();
    }

    public Task BeginClose()
    {
        lock (_stateLock)
        {
            if (_closing)
            {
                return _drained?.Task ?? Task.CompletedTask;
            }

            _closing = true;
            if (_pendingOperations == 0)
            {
                return Task.CompletedTask;
            }

            _drained = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            return _drained.Task;
        }
    }

    public void CancelClose()
    {
        lock (_stateLock)
        {
            _closing = false;
            _drained = null;
        }
    }
}
