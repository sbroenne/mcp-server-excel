using System.Collections.Concurrent;
using System.Text.Json;

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
            ThrowIfStalePackageTransaction(normalizedPath);
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
            ThrowIfStalePackageTransaction(normalizedPath);
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
        var workbookClosed = false;
        try
        {
            await _backend.InvokeAsync(
                "session.close",
                new { filePath = session.FilePath, save },
                session.OperationTimeout);
            workbookClosed = true;
            if (save)
            {
                DeletePackageBaseline(session);
            }
            else
            {
                RestorePackageBaseline(session);
            }
            RemoveSession(sessionId, session);
            return true;
        }
        catch
        {
            if (workbookClosed)
            {
                RemoveSession(sessionId, session);
            }
            else
            {
                session.CancelClose();
            }
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
        if (session.RequiresPackageRecovery)
        {
            throw new InvalidOperationException(
                $"Session '{sessionId}' requires manual Power Query package recovery " +
                "and cannot accept more operations.");
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

    internal async Task MutatePackageAsync(
        MacExcelSession session,
        Action<string> mutateWorkingCopy,
        Func<Task> afterReopen)
    {
        ArgumentNullException.ThrowIfNull(mutateWorkingCopy);
        ArgumentNullException.ThrowIfNull(afterReopen);

        var createdBaseline = session.PackageBaselinePath is null;
        if (createdBaseline)
        {
            var baselinePath = CreateTransactionPath(session.FilePath, "baseline");
            var transactionPath = GetTransactionJournalPath(session.FilePath);
            try
            {
                using (var journal = new FileStream(
                    transactionPath,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None))
                using (var writer = new StreamWriter(journal))
                {
                    writer.Write(JsonSerializer.Serialize(
                        new { baseline = Path.GetFileName(baselinePath) }));
                }
                File.Copy(session.FilePath, baselinePath, overwrite: false);
                session.PackageBaselinePath = baselinePath;
                session.PackageTransactionPath = transactionPath;
            }
            catch
            {
                File.Delete(baselinePath);
                File.Delete(transactionPath);
                throw;
            }
        }

        var checkpointPath = CreateTransactionPath(session.FilePath, "checkpoint");
        var workingPath = CreateTransactionPath(session.FilePath, "working");
        File.Copy(session.FilePath, checkpointPath, overwrite: false);
        File.Copy(session.FilePath, workingPath, overwrite: false);
        var closed = false;
        var reopened = false;
        var preserveCheckpoint = false;
        try
        {
            await _backend.InvokeAsync(
                "session.close",
                new { filePath = session.FilePath, save = false },
                session.OperationTimeout);
            closed = true;

            mutateWorkingCopy(workingPath);
            File.Move(workingPath, session.FilePath, overwrite: true);
            await _backend.InvokeAsync(
                "session.open",
                new { filePath = session.FilePath, show = session.IsVisible },
                session.OperationTimeout);
            reopened = true;
            closed = false;
            await afterReopen();
        }
        catch (Exception operationError)
        {
            if (!closed && !reopened)
            {
                if (createdBaseline)
                {
                    DeletePackageBaseline(session);
                }
                throw;
            }

            try
            {
                if (reopened)
                {
                    await _backend.InvokeAsync(
                        "session.close",
                        new { filePath = session.FilePath, save = false },
                        session.OperationTimeout);
                }

                RestoreFileAtomically(checkpointPath, session.FilePath);
                await _backend.InvokeAsync(
                    "session.open",
                    new { filePath = session.FilePath, show = session.IsVisible },
                    session.OperationTimeout);
            }
            catch (Exception rollbackError)
            {
                preserveCheckpoint = true;
                session.RequiresPackageRecovery = true;
                throw new InvalidOperationException(
                    $"Power Query package mutation failed and automatic rollback also failed. " +
                    $"The recovery checkpoint remains at '{checkpointPath}'.",
                    new AggregateException(operationError, rollbackError));
            }

            if (createdBaseline)
            {
                DeletePackageBaseline(session);
            }

            throw;
        }
        finally
        {
            File.Delete(workingPath);
            if (!preserveCheckpoint && File.Exists(checkpointPath))
            {
                File.Delete(checkpointPath);
            }
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

    private void RemoveSession(string sessionId, MacExcelSession session)
    {
        _sessions.TryRemove(sessionId, out _);
        _paths.TryRemove(session.FilePath, out _);
        session.OperationLock.Dispose();
    }

    private static string CreateTransactionPath(string workbookPath, string role)
    {
        var directory = Path.GetDirectoryName(workbookPath)
            ?? throw new ArgumentException("Workbook path has no parent directory.", nameof(workbookPath));
        return Path.Combine(directory, $".excelmcp-pq-{role}-{Guid.NewGuid():N}.tmp");
    }

    private static string GetTransactionJournalPath(string workbookPath)
    {
        var directory = Path.GetDirectoryName(workbookPath)
            ?? throw new ArgumentException("Workbook path has no parent directory.", nameof(workbookPath));
        return Path.Combine(
            directory,
            $".{Path.GetFileName(workbookPath)}.excelmcp-pq-transaction.json");
    }

    private static void ThrowIfStalePackageTransaction(string workbookPath)
    {
        var transactionPath = GetTransactionJournalPath(workbookPath);
        if (File.Exists(transactionPath))
        {
            throw new InvalidOperationException(
                $"Workbook '{workbookPath}' has an interrupted Power Query package transaction. " +
                $"Inspect the retained transaction record at '{transactionPath}' before reopening.");
        }
    }

    private static void RestorePackageBaseline(MacExcelSession session)
    {
        if (session.PackageBaselinePath is not { } baselinePath)
        {
            return;
        }

        RestoreFileAtomically(baselinePath, session.FilePath);
        DeletePackageBaseline(session);
    }

    private static void RestoreFileAtomically(string sourcePath, string destinationPath)
    {
        var restorePath = CreateTransactionPath(destinationPath, "restore");
        try
        {
            File.Copy(sourcePath, restorePath, overwrite: false);
            File.Move(restorePath, destinationPath, overwrite: true);
        }
        finally
        {
            File.Delete(restorePath);
        }
    }

    private static void DeletePackageBaseline(MacExcelSession session)
    {
        if (session.PackageBaselinePath is not { } baselinePath)
        {
            return;
        }

        File.Delete(baselinePath);
        session.PackageBaselinePath = null;
        if (session.PackageTransactionPath is { } transactionPath)
        {
            File.Delete(transactionPath);
            session.PackageTransactionPath = null;
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
            RestorePackageBaseline(session);
        }
        finally
        {
            _sessions.TryRemove(sessionId, out _);
            _paths.TryRemove(session.FilePath, out _);
            session.OperationLock.Dispose();
        }
    }
}
