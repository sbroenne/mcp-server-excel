using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.Service.Mac;

internal sealed class MacExcelSessionManager : IDisposable
{
    private readonly ConcurrentDictionary<string, MacExcelSession> _sessions = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, string> _paths = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, Task<bool>> _closeTasks = new(StringComparer.Ordinal);
    private readonly MacExcelBackend _backend;
    private readonly Action<string, string> _copyFile;
    private readonly Action<string> _deleteFile;
    private bool _disposed;

    public MacExcelSessionManager(
        MacExcelBackend backend,
        Action<string, string>? copyFile = null,
        Action<string>? deleteFile = null)
    {
        _backend = backend;
        _copyFile = copyFile ?? ((source, destination) =>
            File.Copy(source, destination, overwrite: false));
        _deleteFile = deleteFile ?? File.Delete;
    }

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
            session.RequiresPackageRecovery = true;
            ThrowCloseFailure(closeResult);
        }

        try
        {
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
            RemoveSession(sessionId, session);
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
        var baselinePath = createdBaseline
            ? CreateTransactionPath(session.FilePath, "baseline")
            : null;
        var transactionPath = createdBaseline
            ? GetTransactionJournalPath(session.FilePath)
            : null;
        var checkpointPath = CreateTransactionPath(session.FilePath, "checkpoint");
        var workingPath = CreateTransactionPath(session.FilePath, "working");
        var journalCreated = false;
        var closed = false;
        var reopened = false;
        var setupComplete = false;
        var preserveCheckpoint = false;
        try
        {
            var closeResult = await CloseAndReconcileAsync(
                session,
                "session.close-if-saved",
                new { filePath = session.FilePath });
            if (closeResult.State != WorkbookCloseState.Closed)
            {
                if (closeResult.State == WorkbookCloseState.Indeterminate)
                {
                    session.RequiresPackageRecovery = true;
                }
                ThrowCloseFailure(closeResult);
            }
            closed = true;

            if (createdBaseline)
            {
                using (var journal = new FileStream(
                    transactionPath!,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None))
                {
                    journalCreated = true;
                    using var writer = new StreamWriter(journal);
                    writer.Write(JsonSerializer.Serialize(
                        new { baseline = Path.GetFileName(baselinePath!) }));
                }
                _copyFile(session.FilePath, baselinePath!);
                session.PackageBaselinePath = baselinePath;
                session.PackageTransactionPath = transactionPath;
            }

            _copyFile(session.FilePath, checkpointPath);
            _copyFile(session.FilePath, workingPath);
            setupComplete = true;

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
                throw;
            }

            if (!setupComplete)
            {
                Exception? cleanupError = null;
                Exception? reopenError = null;
                try
                {
                    CleanupNewPackageBaseline(
                        session,
                        createdBaseline,
                        baselinePath,
                        transactionPath,
                        journalCreated);
                }
                catch (Exception error)
                {
                    cleanupError = error;
                }
                try
                {
                    await _backend.InvokeAsync(
                        "session.open",
                        new { filePath = session.FilePath, show = session.IsVisible },
                        session.OperationTimeout);
                    closed = false;
                }
                catch (Exception error)
                {
                    reopenError = error;
                }

                if (cleanupError is not null || reopenError is not null)
                {
                    session.RequiresPackageRecovery = true;
                    var recoveryErrors = new List<Exception> { operationError };
                    if (cleanupError is not null)
                    {
                        recoveryErrors.Add(cleanupError);
                    }
                    if (reopenError is not null)
                    {
                        recoveryErrors.Add(reopenError);
                    }
                    throw new InvalidOperationException(
                        "Power Query package setup failed and automatic cleanup or reopen did not complete.",
                        new AggregateException(recoveryErrors));
                }
                throw;
            }

            try
            {
                if (reopened)
                {
                    var rollbackClose = await CloseAndReconcileAsync(
                        session,
                        "session.close",
                        new { filePath = session.FilePath, save = false });
                    if (rollbackClose.State != WorkbookCloseState.Closed)
                    {
                        ThrowCloseFailure(rollbackClose);
                    }
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

    private void CleanupNewPackageBaseline(
        MacExcelSession session,
        bool createdBaseline,
        string? baselinePath,
        string? transactionPath,
        bool journalCreated)
    {
        if (!createdBaseline)
        {
            return;
        }
        if (session.PackageBaselinePath is not null)
        {
            DeletePackageBaseline(session);
            return;
        }

        if (journalCreated)
        {
            _deleteFile(transactionPath!);
        }
        _deleteFile(baselinePath!);
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
            "the session requires manual package recovery.",
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

    private void RestorePackageBaseline(MacExcelSession session)
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

    private void DeletePackageBaseline(MacExcelSession session)
    {
        if (session.PackageBaselinePath is not { } baselinePath)
        {
            return;
        }

        if (session.PackageTransactionPath is { } transactionPath)
        {
            _deleteFile(transactionPath);
            session.PackageTransactionPath = null;
        }
        _deleteFile(baselinePath);
        session.PackageBaselinePath = null;
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
            var closeResult = await CloseAndReconcileAsync(
                session,
                "session.close",
                new { filePath = session.FilePath, save = false });
            if (closeResult.State != WorkbookCloseState.Closed)
            {
                session.RequiresPackageRecovery = true;
                ThrowCloseFailure(closeResult);
            }
            RestorePackageBaseline(session);
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
