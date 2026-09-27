using System.Collections.Concurrent;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelSessionManager : IDisposable
{
    private readonly ConcurrentDictionary<string, MacExcelSession> _sessions = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, string> _paths = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, Task<bool>> _closeTasks = new(StringComparer.Ordinal);
    private readonly MacExcelBackend _backend;
    private bool _disposed;

    public MacExcelSessionManager(MacExcelBackend backend) => _backend = backend;

    public int Count => _sessions.Count;
    public IReadOnlyCollection<MacExcelSession> Sessions => _sessions.Values.ToArray();

    public async Task<string> CreateAsync(string filePath, bool macroEnabled, bool show, TimeSpan timeout)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var normalizedPath = NormalizeAndClaim(filePath, out var sessionId);
        var packageCreated = false;
        try
        {
            MacWorkbookPackage.Create(normalizedPath, macroEnabled);
            packageCreated = true;
            await _backend.InvokeAsync("session.open", new { filePath = normalizedPath, show }, timeout);
            AddSession(sessionId, normalizedPath, timeout, show);
            return sessionId;
        }
        catch
        {
            _paths.TryRemove(normalizedPath, out _);
            if (packageCreated)
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
        try
        {
            await _backend.InvokeAsync(
                "session.close",
                new { filePath = session.FilePath, save },
                session.OperationTimeout);
            _sessions.TryRemove(sessionId, out _);
            _paths.TryRemove(session.FilePath, out _);
            session.OperationLock.Dispose();
            return true;
        }
        catch
        {
            session.CancelClose();
            throw;
        }
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

        if (!session.TryAdmitOperation())
        {
            throw new KeyNotFoundException($"Session '{sessionId}' is closing.");
        }

        await session.OperationLock.WaitAsync();
        Interlocked.Increment(ref session.ActiveOperations);
        try
        {
            return await operation(session);
        }
        finally
        {
            Interlocked.Decrement(ref session.ActiveOperations);
            session.OperationLock.Release();
            session.CompleteOperation();
        }
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
            await _backend.InvokeAsync(
                "session.close",
                new { filePath = session.FilePath, save = false },
                session.OperationTimeout);
        }
        finally
        {
            _sessions.TryRemove(sessionId, out _);
            _paths.TryRemove(session.FilePath, out _);
            session.OperationLock.Dispose();
        }
    }
}
