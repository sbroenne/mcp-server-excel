using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelSessionManager : IDisposable
{
    private readonly ConcurrentDictionary<string, MacExcelSession> _sessions = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, string> _paths = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, Task<bool>> _closeTasks = new(StringComparer.Ordinal);
    private readonly MacExcelBackend _backend;
    private readonly Action<string, bool> _createWorkbook;
    private bool _disposed;

    public MacExcelSessionManager(
        MacExcelBackend backend,
        Action<string, bool>? createWorkbook = null)
    {
        _backend = backend;
        _createWorkbook = createWorkbook ?? MacWorkbookTemplate.Copy;
    }

    public int Count => _sessions.Count;
    public IReadOnlyCollection<MacExcelSession> Sessions => _sessions.Values.ToArray();

    public async Task<string> CreateAsync(string filePath, bool macroEnabled, bool show, TimeSpan timeout)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var normalizedPath = NormalizeAndClaim(filePath, out var sessionId);
        var workbookCreated = false;
        try
        {
            _createWorkbook(normalizedPath, macroEnabled);
            workbookCreated = true;
            await _backend.InvokeAsync("session.open", new { filePath = normalizedPath, show }, timeout);
            AddSession(sessionId, normalizedPath, timeout, show);
            return sessionId;
        }
        catch (Exception error)
        {
            _paths.TryRemove(normalizedPath, out _);
            if (workbookCreated && error is not MacExcelOperationException { ErrorCategory: "RecoveryRequired" })
            {
                File.Delete(normalizedPath);
            }
            throw;
        }
    }

    public async Task<string> OpenAsync(string filePath, bool show, TimeSpan timeout)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var normalizedPath = NormalizeAndClaim(filePath, out var sessionId);
        try
        {
            await _backend.InvokeAsync("session.open", new { filePath = normalizedPath, show }, timeout);
            AddSession(sessionId, normalizedPath, timeout, show);
            return sessionId;
        }
        catch
        {
            _paths.TryRemove(normalizedPath, out _);
            throw;
        }
    }

    public async Task<bool> CloseAsync(string sessionId, bool save)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (!_sessions.TryGetValue(sessionId, out var session))
        {
            return false;
        }

        var closeTask = _closeTasks.GetOrAdd(
            sessionId,
            _ => CloseSessionAsync(sessionId, session, save));
        try
        {
            return await closeTask;
        }
        finally
        {
            _closeTasks.TryRemove(new KeyValuePair<string, Task<bool>>(sessionId, closeTask));
        }
    }

    private async Task<bool> CloseSessionAsync(string sessionId, MacExcelSession session, bool save)
    {
        await session.BeginClose();
        if (session.HasUnconfirmedOpen)
        {
            throw new MacExcelOperationException(
                "RecoveryRequired",
                "The workbook has an unconfirmed file-open request. Automatic close or rollback is unsafe. " +
                "Resolve pending Excel dialogs and reconcile the exact workbook and retained recovery files " +
                "before restarting the ExcelMcp client.");
        }
        var closeResult = await CloseAndReconcileAsync(
            session,
            "session.close",
            new { filePath = session.FilePath, save });
        if (closeResult.State == WorkbookCloseState.Open)
        {
            session.CancelClose();
            ThrowCloseFailure(closeResult);
        }
        if (closeResult.State == WorkbookCloseState.Indeterminate)
        {
            session.RequiresRecovery = true;
            ThrowCloseFailure(closeResult);
        }

        RemoveSession(sessionId, session);
        return true;
    }

    public async Task<T> ExecuteAsync<T>(
        string sessionId,
        Func<MacExcelSession, Task<T>> operation)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (!_sessions.TryGetValue(sessionId, out var session))
        {
            throw new KeyNotFoundException($"Session '{sessionId}' not found.");
        }
        if (session.RequiresRecovery)
        {
            throw new InvalidOperationException(
                $"Session '{sessionId}' requires manual recovery " +
                "and cannot accept more operations.");
        }
        if (session.UnsafeReason is not null)
        {
            throw new InvalidOperationException(
                $"Session '{sessionId}' is unsafe after an uncertain Office.js mutation: " +
                session.UnsafeReason);
        }

        if (!session.TryAdmitOperation())
        {
            throw new KeyNotFoundException($"Session '{sessionId}' is closing.");
        }

        await session.OperationLock.WaitAsync();
        Interlocked.Increment(ref session.ActiveOperations);
        try
        {
            if (session.RequiresRecovery)
            {
                throw new InvalidOperationException(
                    $"Session '{sessionId}' requires manual recovery " +
                    "and cannot accept more operations.");
            }
            if (session.UnsafeReason is not null)
            {
                throw new InvalidOperationException(
                    $"Session '{sessionId}' is unsafe after an uncertain Office.js mutation: " +
                    session.UnsafeReason);
            }
            return await operation(session);
        }
        finally
        {
            Interlocked.Decrement(ref session.ActiveOperations);
            session.OperationLock.Release();
            session.CompleteOperation();
        }
    }

    internal void RequireRecovery(string sessionId)
    {
        if (_sessions.TryGetValue(sessionId, out var session))
        {
            session.RequiresRecovery = true;
        }
    }

    private async Task<WorkbookCloseResult> CloseAndReconcileAsync(
        MacExcelSession session,
        string command,
        object arguments)
    {
        try
        {
            await _backend.InvokeAsync(command, arguments, session.OperationTimeout);
            return new WorkbookCloseResult(WorkbookCloseState.Closed, null);
        }
        catch (Exception closeError)
        {
            try
            {
                var state = await _backend.InvokeAsync(
                    "session.is-open",
                    new { filePath = session.FilePath },
                    session.OperationTimeout);
                return new WorkbookCloseResult(
                    state.GetProperty("open").GetBoolean()
                        ? WorkbookCloseState.Open
                        : WorkbookCloseState.Closed,
                    closeError);
            }
            catch (Exception stateError)
            {
                return new WorkbookCloseResult(
                    WorkbookCloseState.Indeterminate,
                    new AggregateException(closeError, stateError));
            }
        }
    }

    [DoesNotReturn]
    private static void ThrowCloseFailure(WorkbookCloseResult result)
    {
        if (result.State == WorkbookCloseState.Open && result.Error is not null)
        {
            ExceptionDispatchInfo.Capture(result.Error).Throw();
        }

        throw new InvalidOperationException(
            "The service could not determine the exact workbook close outcome; " +
            "the session requires manual workbook recovery.",
            result.Error);
    }

    private string NormalizeAndClaim(string filePath, out string sessionId)
    {
        var normalizedPath = Path.GetFullPath(filePath);
        sessionId = Guid.NewGuid().ToString("N");
        if (!_paths.TryAdd(normalizedPath, sessionId))
        {
            throw new InvalidOperationException(
                $"File '{normalizedPath}' is already open in another session.");
        }

        return normalizedPath;
    }

    private void AddSession(string sessionId, string filePath, TimeSpan timeout, bool show)
    {
        var session = new MacExcelSession
        {
            SessionId = sessionId,
            FilePath = filePath,
            OperationTimeout = timeout,
            IsVisible = show
        };
        if (!_sessions.TryAdd(sessionId, session))
        {
            session.OperationLock.Dispose();
            throw new InvalidOperationException($"Session ID collision: {sessionId}");
        }
    }

    private void RemoveSession(string sessionId, MacExcelSession session)
    {
        _sessions.TryRemove(sessionId, out _);
        _paths.TryRemove(session.FilePath, out _);
        session.OperationLock.Dispose();
    }


    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        foreach (var sessionId in _sessions.Keys)
        {
            try
            {
                CloseForDisposeAsync(sessionId).GetAwaiter().GetResult();
            }
            catch
            {
                // The process does not own shared Excel and must never kill it during disposal.
            }
        }
        _sessions.Clear();
        _paths.Clear();
        _closeTasks.Clear();
    }

    private async Task CloseForDisposeAsync(string sessionId)
    {
        if (!_sessions.TryGetValue(sessionId, out var session))
        {
            return;
        }

        await session.BeginClose();
        try
        {
            // A queued open may complete later even when a current closed-state probe returns false.
            if (session.HasUnconfirmedOpen) return;
            var closeResult = await CloseAndReconcileAsync(
                session,
                "session.close",
                new { filePath = session.FilePath, save = false });
            if (closeResult.State != WorkbookCloseState.Closed)
            {
                session.RequiresRecovery = true;
                ThrowCloseFailure(closeResult);
            }
        }
        finally
        {
            _sessions.TryRemove(sessionId, out _);
            _paths.TryRemove(session.FilePath, out _);
            session.OperationLock.Dispose();
        }
    }

    private enum WorkbookCloseState
    {
        Closed,
        Open,
        Indeterminate
    }

    private sealed record WorkbookCloseResult(
        WorkbookCloseState State,
        Exception? Error);
}
