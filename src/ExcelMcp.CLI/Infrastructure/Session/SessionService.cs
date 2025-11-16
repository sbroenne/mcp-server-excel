using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.CLI.Infrastructure.Session;

internal sealed class SessionService : ISessionService, IDisposable
{
    private readonly SessionManager _sessionManager = new();
    private bool _disposed;

    public string Create(string filePath)
    {
        EnsureNotDisposed();
        return _sessionManager.CreateSession(filePath);
    }

    public bool Save(string sessionId)
    {
        EnsureNotDisposed();
        return _sessionManager.SaveSession(sessionId);
    }

    public bool Close(string sessionId)
    {
        EnsureNotDisposed();
        return _sessionManager.CloseSession(sessionId);
    }

    public IReadOnlyList<SessionDescriptor> List()
    {
        EnsureNotDisposed();
        return _sessionManager.GetActiveSessions();
    }

    public IExcelBatch GetBatch(string sessionId)
    {
        EnsureNotDisposed();
        var batch = _sessionManager.GetSession(sessionId);
        if (batch == null)
        {
            throw new InvalidOperationException($"Session '{sessionId}' not found.");
        }

        return batch;
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _sessionManager.Dispose();
        _disposed = true;
    }

    private void EnsureNotDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);
}
