namespace Sbroenne.ExcelMcp.ComInterop.Session;

internal sealed class ExcelApplicationLease : IDisposable
{
    private static readonly Lock RegistryLock = new();
    private static readonly Dictionary<ExcelProcessIdentity, SharedLease> Leases = [];
    private readonly ExcelProcessIdentity _identity;
    private readonly SharedLease _shared;
    private bool _disposed;

    internal ExcelApplicationLease(ExcelProcessIdentity identity)
    {
        _identity = identity;
        lock (RegistryLock)
        {
            if (!Leases.TryGetValue(identity, out var shared))
            {
                shared = new SharedLease(identity);
                Leases.Add(identity, shared);
            }
            _shared = shared;
            _shared.Connections++;
        }
    }

    internal IDisposable EnterExecution() => new ExecutionScope(_shared.SyncRoot);
    internal bool IsExcelVisible => _shared.IsExcelVisible;
    internal long VisibilityRevision => Interlocked.Read(ref _shared.VisibilityRevision);
    internal void UpdateVisibility(bool visible, long? expectedRevision = null)
    {
        lock (_shared.WindowControlLock)
        {
            if (expectedRevision is not null && VisibilityRevision != expectedRevision)
                return;
            _shared.IsExcelVisible = visible;
            Interlocked.Increment(ref _shared.VisibilityRevision);
        }
    }
    internal Lock WindowControlLock => _shared.WindowControlLock;

    public void Dispose()
    {
        lock (RegistryLock)
        {
            if (_disposed)
                return;

            if (_shared.Connections == 1)
            {
                _shared.Dispose();
                Leases.Remove(_identity);
            }
            _shared.Connections--;
            _disposed = true;
        }
    }

    private sealed class ExecutionScope : IDisposable
    {
        private readonly object _syncRoot;

        internal ExecutionScope(object syncRoot)
        {
            _syncRoot = syncRoot;
            Monitor.Enter(_syncRoot);
        }

        public void Dispose() => Monitor.Exit(_syncRoot);
    }

    // A mutex belongs to a thread. Keep its owner alive until the last local
    // connection detaches, even when the original workbook's STA has exited.
    private sealed class SharedLease : IDisposable
    {
        private static readonly TimeSpan WaitTimeout = TimeSpan.FromSeconds(2);
        private readonly ManualResetEventSlim _release = new();
        private readonly TaskCompletionSource _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource _completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly Thread _thread;

        internal object SyncRoot { get; } = new();
        internal Lock WindowControlLock { get; } = new();
        internal volatile bool IsExcelVisible;
        internal long VisibilityRevision;
        internal int Connections { get; set; }

        internal SharedLease(ExcelProcessIdentity identity)
        {
            _thread = new Thread(() => Run(identity))
            {
                IsBackground = true,
                Name = $"ExcelApplicationLease-{identity.ProcessId}"
            };
            _thread.Start();
            try
            {
                _ready.Task.WaitAsync(WaitTimeout).GetAwaiter().GetResult();
            }
            catch (Exception startupFailure)
            {
                _release.Set();
                if (!_thread.Join(WaitTimeout))
                    throw new AggregateException(startupFailure,
                        new TimeoutException("The Excel application lock thread did not finish after startup failed."));
                _ = _completion.Task.Exception;
                _release.Dispose();
                throw;
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "The lock-owner thread transfers startup and release failures to the waiting connection through task completions; failures are never treated as a successful lease.")]
        private void Run(ExcelProcessIdentity identity)
        {
            try
            {
                using var mutex = new Mutex(false,
                    $@"Local\ExcelMcp.Application.{identity.ProcessId}.{identity.StartedAtUtcFileTime}");
                bool acquired;
                try
                {
                    acquired = mutex.WaitOne(TimeSpan.Zero);
                }
                catch (AbandonedMutexException)
                {
                    acquired = true;
                }
                if (!acquired)
                    throw new InvalidOperationException(
                        "This Excel application is already managed by another active ExcelMcp host.");
                try
                {
                    _ready.TrySetResult();
                    _release.Wait();
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
                _completion.TrySetResult();
            }
            catch (Exception ex)
            {
                _ready.TrySetException(ex);
                _completion.TrySetException(ex);
            }
        }

        public void Dispose()
        {
            _release.Set();
            _completion.Task.WaitAsync(WaitTimeout).GetAwaiter().GetResult();
            if (!_thread.Join(WaitTimeout))
                throw new TimeoutException("The Excel application lock thread did not finish after release.");
            _release.Dispose();
        }
    }
}
