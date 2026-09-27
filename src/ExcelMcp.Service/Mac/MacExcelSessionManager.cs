using System.Collections.Concurrent;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelSessionManager : IDisposable
{
    private readonly ConcurrentDictionary<string, MacExcelSession> _sessions = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, string> _paths = new(StringComparer.Ordinal);
    private readonly MacExcelBackend _backend;
    private bool _disposed;

    public MacExcelSessionManager(MacExcelBackend backend) => _backend = backend;

    public int Count => _sessions.Count;
    public IReadOnlyCollection<MacExcelSession> Sessions => _sessions.Values.ToArray();

    public async Task<string> CreateAsync(string filePath, bool macroEnabled, bool show, TimeSpan timeout)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var normalizedPath = NormalizeAndClaim(filePath, out var sessionId);
        try
        {
            await _backend.InvokeAsync(
                "session.create",
                new { filePath = normalizedPath, macroEnabled, show },
                timeout);
            AddSession(sessionId, normalizedPath, timeout, show);
            return sessionId;
        }
        catch
        {
            _paths.TryRemove(normalizedPath, out _);
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
        if (!_sessions.TryGetValue(sessionId, out var session))
        {
            return false;
        }

        await session.OperationLock.WaitAsync();
        try
        {
            if (Volatile.Read(ref session.ActiveOperations) != 0)
            {
                throw new InvalidOperationException(
                    $"Session '{sessionId}' has active operations and cannot be closed.");
            }

            await _backend.InvokeAsync(
                "session.close",
                new { filePath = session.FilePath, save },
                session.OperationTimeout);
            _sessions.TryRemove(sessionId, out _);
            _paths.TryRemove(session.FilePath, out _);
            session.OperationLock.Dispose();
            return true;
        }
        finally
        {
            if (_sessions.ContainsKey(sessionId))
            {
                session.OperationLock.Release();
            }
        }
    }

    public async Task<T> ExecuteAsync<T>(
        string sessionId,
        Func<MacExcelSession, Task<T>> operation)
    {
        if (!_sessions.TryGetValue(sessionId, out var session))
        {
            throw new KeyNotFoundException($"Session '{sessionId}' not found.");
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
        foreach (var session in _sessions.Values)
        {
            try
            {
                _backend.InvokeAsync(
                    "session.close",
                    new { filePath = session.FilePath, save = false },
                    session.OperationTimeout).GetAwaiter().GetResult();
            }
            catch
            {
                // The process does not own shared Excel and must never kill it during disposal.
            }
            session.OperationLock.Dispose();
        }
        _sessions.Clear();
        _paths.Clear();
    }
}
