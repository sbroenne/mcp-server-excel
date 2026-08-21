using System.Collections.Concurrent;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Manages active Excel sessions for the MCP server and CLI.
/// Maps user-facing sessionId to internal IExcelBatch instances.
/// </summary>
/// <remarks>
/// <para><b>Concurrency Model:</b></para>
/// <list type="bullet">
/// <item><b>Within-session operations are SERIAL:</b> Each session queues operations on one STA thread</item>
/// <item><b>Between-session operations CAN be parallel:</b> Different sessions = different Excel processes</item>
/// <item><b>Same-file prevention:</b> Cannot open the same file in multiple sessions (matches Excel UI behavior)</item>
/// </list>
/// <para><b>Resource Limits:</b></para>
/// <list type="bullet">
/// <item>Each session = one Excel.Application process (~50-100MB+ memory)</item>
/// <item>Recommended maximum: 3-5 concurrent sessions on typical desktop machines</item>
/// <item>Always close sessions promptly to free resources</item>
/// </list>
/// </remarks>
public sealed class SessionManager : IDisposable
{
    private static readonly ConcurrentDictionary<ExcelProcessIdentity, byte> _trackedExcelProcesses = new();
    private static int _processExitRegistered;

    /// <summary>
    /// Raised whenever the set of Excel processes owned by this process changes.
    /// </summary>
    /// <remarks>Subscribers are notified asynchronously and cannot interrupt session lifecycle operations.</remarks>
    public static event Action<IReadOnlyCollection<int>>? TrackedExcelProcessesChanged;

    /// <summary>
    /// Raised with PID/start-time identities whenever owned Excel processes change.
    /// </summary>
    public static event Action<IReadOnlyCollection<ExcelProcessIdentity>>? TrackedExcelProcessIdentitiesChanged;

    /// <summary>
    /// Invoked synchronously when a new identity is first tracked so in-process
    /// owners can durably persist it before session startup continues.
    /// </summary>
    internal static event Action<ExcelProcessIdentity>? ExcelProcessIdentityTracked;

    /// <summary>
    /// Registers an Excel process ID for cleanup on unexpected process exit.
    /// Called from ExcelBatch when a PID is captured.
    /// </summary>
    public static void TrackExcelProcess(int processId) =>
        _ = TrackExcelProcessIdentity(processId);

    /// <summary>
    /// Registers an Excel process and returns its captured PID/start-time identity.
    /// </summary>
    public static ExcelProcessIdentity? TrackExcelProcessIdentity(int processId)
    {
        if (!TryCaptureExcelProcessIdentity(processId, out var identity))
        {
            return null;
        }

        TrackExcelProcess(identity);
        return identity;
    }

    internal static void TrackExcelProcess(ExcelProcessIdentity identity)
    {
        _trackedExcelProcesses[identity] = 0;

        // Register handler exactly once (thread-safe)
        if (Interlocked.CompareExchange(ref _processExitRegistered, 1, 0) == 0)
        {
            AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
        }

        NotifyExcelProcessIdentityTracked(identity);
        NotifyTrackedExcelProcessesChanged();
    }

    /// <summary>
    /// Marks an Excel process as no longer needing cleanup.
    /// </summary>
    public static void UntrackExcelProcess(int processId)
    {
        foreach (var identity in _trackedExcelProcesses.Keys
                     .Where(process => process.ProcessId == processId))
        {
            _trackedExcelProcesses.TryRemove(identity, out _);
        }

        NotifyTrackedExcelProcessesChanged();
    }

    /// <summary>
    /// Marks one exact PID/start-time identity as no longer needing cleanup.
    /// </summary>
    public static void UntrackExcelProcess(ExcelProcessIdentity identity)
    {
        _trackedExcelProcesses.TryRemove(identity, out _);
        NotifyTrackedExcelProcessesChanged();
    }

    /// <summary>
    /// Returns a snapshot of Excel process identities currently owned by this process.
    /// </summary>
    public static IReadOnlyCollection<ExcelProcessIdentity> GetTrackedExcelProcesses() =>
        _trackedExcelProcesses.Keys.ToArray();

    /// <summary>
    /// Returns the PIDs currently tracked for backward compatibility.
    /// </summary>
    public static IReadOnlyCollection<int> GetTrackedExcelProcessIds() =>
        _trackedExcelProcesses.Keys
            .Select(process => process.ProcessId)
            .Distinct()
            .ToArray();

    private static void NotifyTrackedExcelProcessesChanged()
    {
        var legacySubscribers = TrackedExcelProcessesChanged;
        var identitySubscribers = TrackedExcelProcessIdentitiesChanged;
        if (legacySubscribers is null && identitySubscribers is null)
        {
            return;
        }

        var processes = GetTrackedExcelProcesses();
        var processIds = processes
            .Select(process => process.ProcessId)
            .Distinct()
            .ToArray();
        _ = Task.Run(() =>
        {
            foreach (var subscriber in legacySubscribers?.GetInvocationList() ?? [])
            {
                try
                {
                    ((Action<IReadOnlyCollection<int>>)subscriber)(processIds);
                }
                catch (Exception ex)
                {
                    Trace.TraceWarning(
                        "Tracked Excel process notification failed: {0}",
                        ex.Message);
                }
            }

            foreach (var subscriber in identitySubscribers?.GetInvocationList() ?? [])
            {
                try
                {
                    ((Action<IReadOnlyCollection<ExcelProcessIdentity>>)subscriber)(processes);
                }
                catch (Exception ex)
                {
                    Trace.TraceWarning(
                        "Tracked Excel process identity notification failed: {0}",
                        ex.Message);
                }
            }
        });
    }

    private static void NotifyExcelProcessIdentityTracked(ExcelProcessIdentity identity)
    {
        var subscribers = ExcelProcessIdentityTracked;
        if (subscribers is null)
        {
            return;
        }

        foreach (var subscriber in subscribers.GetInvocationList())
        {
            try
            {
                ((Action<ExcelProcessIdentity>)subscriber)(identity);
            }
            catch (Exception ex)
            {
                throw new ExcelProcessPersistenceException(identity, ex);
            }
        }
    }

    private static void OnProcessExit(object? sender, EventArgs e)
    {
        // INSTRUMENTATION: Track ProcessExit handler execution
        int killedCount = 0;
        int alreadyExitedCount = 0;
        int failedCount = 0;

        foreach (var identity in _trackedExcelProcesses.Keys)
        {
            try
            {
                var stopped = OwnedProcessGuard.TryTerminate(
                        identity,
                        TimeSpan.Zero,
                        TimeSpan.FromSeconds(5),
                        out var terminated);
                if (stopped && terminated)
                {
                    killedCount++;
                    // Note: Cannot use ILogger here - ProcessExit handler runs during AppDomain teardown
                    SessionDiagnostics.WriteStdErr($"[DIAG-PROCESSEXIT-KILLED] Force-killed Excel process {identity.ProcessId}");
                }
                else if (stopped)
                {
                    alreadyExitedCount++;
                }
                else
                {
                    failedCount++;
                    SessionDiagnostics.WriteStdErr($"[DIAG-PROCESSEXIT-FAILED] Failed to kill Excel process {identity.ProcessId}");
                }
            }
            catch (ArgumentException)
            {
                // Process already exited
                alreadyExitedCount++;
            }
            catch (Exception ex)
            {
                // Process inaccessible
                failedCount++;
                SessionDiagnostics.WriteStdErr($"[DIAG-PROCESSEXIT-FAILED] Failed to kill Excel process {identity.ProcessId}: {ex.Message}");
            }
        }

        // Summary log
        if (killedCount > 0 || failedCount > 0)
        {
            SessionDiagnostics.WriteStdErr($"[DIAG-PROCESSEXIT-SUMMARY] Killed={killedCount}, AlreadyExited={alreadyExitedCount}, Failed={failedCount}, Total={_trackedExcelProcesses.Count}");
        }
    }

    private static bool TryCaptureExcelProcessIdentity(
        int processId,
        out ExcelProcessIdentity identity)
    {
        identity = default;
        try
        {
            using var process = Process.GetProcessById(processId);
            if (process.HasExited)
            {
                return false;
            }

            identity = new ExcelProcessIdentity(
                processId,
                process.StartTime.ToUniversalTime().ToFileTimeUtc());
            return true;
        }
        catch (ArgumentException ex)
        {
            Trace.TraceWarning("Could not track Excel process {0}: {1}", processId, ex.Message);
            return false;
        }
        catch (InvalidOperationException ex)
        {
            Trace.TraceWarning("Could not track Excel process {0}: {1}", processId, ex.Message);
            return false;
        }
        catch (Win32Exception ex)
        {
            Trace.TraceWarning("Could not track Excel process {0}: {1}", processId, ex.Message);
            return false;
        }
    }

    private readonly ConcurrentDictionary<string, IExcelBatch> _activeSessions = new();
    private readonly ConcurrentDictionary<string, string> _activeFilePaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, string> _sessionFilePaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, int> _activeOperationCounts = new();
    private readonly ConcurrentDictionary<string, object> _sessionLocks = new();
    private readonly ConcurrentDictionary<string, byte> _closingSessions = new();
    private readonly ConcurrentDictionary<string, ExceptionDispatchInfo> _teardownFailures = new();
    private readonly ConcurrentDictionary<string, bool> _showExcelFlags = new();
    private readonly ConcurrentDictionary<string, SessionOrigin> _sessionOrigins = new();
    private readonly ConcurrentDictionary<string, DateTime> _sessionCreatedAt = new();
    private readonly ConcurrentDictionary<string, string> _sessionFilePathReservations = new();
    private readonly object _filePathReservationLock = new();
    private readonly Polly.ResiliencePipeline _sessionCreationPipeline = ResiliencePipelines.CreateSessionCreationPipeline();
    private readonly ILogger<SessionManager> _logger;
    private bool _disposed;

    private object GetSessionLock(string sessionId) =>
        _sessionLocks.GetOrAdd(sessionId, static _ => new object());

    private bool TryClaimFilePath(string normalizedPath, string sessionId)
    {
        lock (_filePathReservationLock)
        {
            return _activeFilePaths.TryAdd(normalizedPath, sessionId);
        }
    }

    private void ReleaseFilePathClaim(string normalizedPath, string sessionId)
    {
        lock (_filePathReservationLock)
        {
            if (_activeFilePaths.TryGetValue(normalizedPath, out var ownerSessionId) &&
                string.Equals(ownerSessionId, sessionId, StringComparison.Ordinal))
            {
                _activeFilePaths.TryRemove(normalizedPath, out _);
            }
        }
    }

    /// <summary>
    /// Creates a new SessionManager with optional logging.
    /// </summary>
    /// <param name="logger">Optional logger for diagnostics</param>
    public SessionManager(ILogger<SessionManager>? logger = null)
    {
        _logger = logger ?? NullLogger<SessionManager>.Instance;
    }

    /// <summary>
    /// Creates a new session for the specified Excel file.
    /// </summary>
    /// <param name="filePath">Path to the Excel file to open</param>
    /// <param name="show">Whether to show the Excel window (default: false for background automation)</param>
    /// <param name="operationTimeout">Maximum time for startup and any operation in this session (default: 120 seconds)</param>
    /// <param name="origin">Which client is creating this session (CLI or MCP)</param>
    /// <returns>Unique session ID for this session</returns>
    /// <exception cref="FileNotFoundException">File does not exist</exception>
    /// <exception cref="InvalidOperationException">Failed to create session or file already open in another session</exception>
    /// <remarks>
    /// <para><b>Resource Impact:</b> Creates a new Excel.Application process (~50-100MB+ memory).</para>
    /// <para><b>Same-file prevention:</b> Throws if file is already open in another session.</para>
    /// <para><b>Concurrency:</b> You can create multiple sessions for DIFFERENT files. Operations within each session execute serially.</para>
    /// </remarks>
    public string CreateSession(string filePath, bool show = false, TimeSpan? operationTimeout = null, SessionOrigin origin = SessionOrigin.Unknown)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Excel file not found: {filePath}. To create a new file, use the 'create' action instead of 'open'.", filePath);
        }

        // Normalize file path for comparison
        string normalizedPath = Path.GetFullPath(filePath);

        // Generate unique session ID
        string sessionId = Guid.NewGuid().ToString("N");
        if (!TryClaimFilePath(normalizedPath, sessionId))
        {
            throw new InvalidOperationException(
                $"File '{filePath}' is already open in another session or reserved for one. " +
                "Excel cannot open the same file multiple times.");
        }

        IExcelBatch? batch = null;
        try
        {
            // Reject external file-access failures after the in-process path claim,
            // so a duplicate session reports its canonical ownership error instead
            // of being mistaken for an unrelated OS-level file lock.
            if (!FileAccessValidator.IsIrmProtected(normalizedPath))
            {
                FileAccessValidator.ValidateFileNotLocked(normalizedPath);
            }

            // Create batch session using Core API with retry for transient COM failures
            // (e.g., CO_E_SERVER_EXEC_FAILURE when system resources are constrained)
            batch = _sessionCreationPipeline.Execute(() => ExcelSession.BeginBatch(show, operationTimeout, filePath));

            // Store in active sessions
            if (!_activeSessions.TryAdd(sessionId, batch))
            {
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
            }

            if (!_sessionFilePaths.TryAdd(sessionId, normalizedPath))
            {
                _activeSessions.TryRemove(sessionId, out _);
                throw new InvalidOperationException($"Failed to record session metadata for: {sessionId}");
            }

            // Initialize operation counter and show flag
            _activeOperationCounts[sessionId] = 0;
            _showExcelFlags[sessionId] = show;
            _sessionOrigins[sessionId] = origin;
            _sessionCreatedAt[sessionId] = DateTime.UtcNow;

            // Success - transfer ownership to dictionary
            var result = sessionId;
            batch = null;  // Prevent disposal in finally
            return result;
        }
        catch (Exception ex)
        {
            _activeSessions.TryRemove(sessionId, out _);
            _sessionFilePaths.TryRemove(sessionId, out _);
            ReleaseFilePathClaim(normalizedPath, sessionId);
            throw new InvalidOperationException($"Failed to create session for '{filePath}': {ex.Message}", ex);
        }
        finally
        {
            // Dispose batch only if we didn't successfully add it to dictionary
            batch?.Dispose();
        }
    }

    /// <summary>
    /// Creates a new Excel file and opens a session for it in one operation.
    /// This is the preferred method for creating new workbooks with sessions.
    /// </summary>
    /// <param name="filePath">Path for the new Excel file (.xlsx or .xlsm)</param>
    /// <param name="show">Whether to show the Excel window (default: false)</param>
    /// <param name="operationTimeout">Maximum time for startup and any operation in this session (default: 120 seconds)</param>
    /// <param name="origin">Which client is creating this session (CLI or MCP)</param>
    /// <returns>Unique session ID for this session</returns>
    /// <exception cref="InvalidOperationException">File already exists, or failed to create session</exception>
    /// <exception cref="DirectoryNotFoundException">Target directory does not exist</exception>
    /// <remarks>
    /// <para><b>Single Excel Start:</b> This method starts Excel only once, creating the file and session together.</para>
    /// <para><b>File Format:</b> Determined by extension - .xlsm creates macro-enabled workbook.</para>
    /// <para><b>Directory:</b> Target directory must exist - will not be created automatically.</para>
    /// </remarks>
    public string CreateSessionForNewFile(string filePath, bool show = false, TimeSpan? operationTimeout = null, SessionOrigin origin = SessionOrigin.Unknown)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        string normalizedPath = Path.GetFullPath(filePath);

        string? directory = Path.GetDirectoryName(normalizedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException(
                $"Directory does not exist: '{directory}'. Create the directory first before creating Excel files.");
        }

        // Validate extension
        string extension = Path.GetExtension(normalizedPath).ToLowerInvariant();
        if (extension is not (".xlsx" or ".xlsm"))
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only .xlsx and .xlsm are supported.");
        }

        // Check if file already exists
        if (File.Exists(normalizedPath))
        {
            throw new InvalidOperationException($"File already exists: {normalizedPath}. Use CreateSession to open existing files.");
        }

        // Generate unique session ID
        string sessionId = Guid.NewGuid().ToString("N");
        if (!TryClaimFilePath(normalizedPath, sessionId))
        {
            throw new InvalidOperationException($"File '{filePath}' is already open or reserved by another session.");
        }

        bool isMacroEnabled = extension == ".xlsm";

        ExcelBatch? batch = null;
        try
        {
            // Create new workbook and keep session open with retry for transient COM failures
            batch = _sessionCreationPipeline.Execute(() => ExcelBatch.CreateNewWorkbook(normalizedPath, isMacroEnabled, logger: null, show: show, operationTimeout: operationTimeout));

            // Store in active sessions
            if (!_activeSessions.TryAdd(sessionId, batch))
            {
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
            }

            if (!_sessionFilePaths.TryAdd(sessionId, normalizedPath))
            {
                _activeSessions.TryRemove(sessionId, out _);
                throw new InvalidOperationException($"Failed to record session metadata for: {sessionId}");
            }

            // Initialize operation counter and show flag
            _activeOperationCounts[sessionId] = 0;
            _showExcelFlags[sessionId] = show;
            _sessionOrigins[sessionId] = origin;
            _sessionCreatedAt[sessionId] = DateTime.UtcNow;

            // Success - transfer ownership to dictionary
            var result = sessionId;
            batch = null;  // Prevent disposal in finally
            return result;
        }
        catch (Exception ex)
        {
            _activeSessions.TryRemove(sessionId, out _);
            _sessionFilePaths.TryRemove(sessionId, out _);
            ReleaseFilePathClaim(normalizedPath, sessionId);
            throw new InvalidOperationException($"Failed to create session for new file '{filePath}': {ex.Message}", ex);
        }
        finally
        {
            // Dispose batch only if we didn't successfully add it to dictionary
            batch?.Dispose();
        }
    }



    /// <summary>
    /// Gets an active session by ID.
    /// If the session exists but Excel has died, it is automatically cleaned up and null is returned.
    /// </summary>
    /// <param name="sessionId">Session ID returned from CreateSession</param>
    /// <returns>IExcelBatch instance, or null if session not found or Excel process is dead</returns>
    public IExcelBatch? GetSession(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return null;
        }

        var sessionLock = GetSessionLock(sessionId);
        lock (sessionLock)
        {
            if (!_activeSessions.TryGetValue(sessionId, out var batch))
            {
                return null;
            }

            // Check if Excel process is still alive
            if (!batch.IsExcelProcessAlive())
            {
                _logger?.LogWarning("Session {SessionId} has dead Excel process, auto-cleaning up", sessionId);
                CleanupDeadSession(sessionId, batch);
                return null;
            }

            return batch;
        }
    }

    /// <summary>
    /// Cleans up a session whose Excel process has died.
    /// This removes all tracking data and disposes the batch (best effort).
    /// </summary>
    private void CleanupDeadSession(string sessionId, IExcelBatch batch)
    {
        if (_teardownFailures.ContainsKey(sessionId)
            && !TryConfirmFailedTeardown(batch))
        {
            return;
        }

        RemoveSessionTracking(sessionId, removeSessionLock: true);

        // Dispose the batch (best effort - process is already dead)
        try
        {
            batch.Dispose();
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Error disposing dead session {SessionId} (expected - process is dead)", sessionId);
        }
    }

    private void RemoveSessionTracking(string sessionId, bool removeSessionLock)
    {
        _activeSessions.TryRemove(sessionId, out _);

        lock (_filePathReservationLock)
        {
            _sessionFilePaths.TryRemove(sessionId, out _);
            foreach (var filePath in _activeFilePaths
                         .Where(kvp => string.Equals(kvp.Value, sessionId, StringComparison.Ordinal))
                         .Select(kvp => kvp.Key)
                         .ToList())
            {
                _activeFilePaths.TryRemove(filePath, out _);
            }

            _sessionFilePathReservations.TryRemove(sessionId, out _);
        }

        _activeOperationCounts.TryRemove(sessionId, out _);
        _closingSessions.TryRemove(sessionId, out _);
        _showExcelFlags.TryRemove(sessionId, out _);
        _sessionOrigins.TryRemove(sessionId, out _);
        _sessionCreatedAt.TryRemove(sessionId, out _);
        _teardownFailures.TryRemove(sessionId, out _);

        if (removeSessionLock)
        {
            _sessionLocks.TryRemove(sessionId, out _);
        }
    }

    /// <summary>
    /// Increments the active operation count for a session.
    /// Call this when starting an operation on the session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    public void BeginOperation(string sessionId)
    {
        if (!TryBeginOperation(sessionId, out _, out var errorMessage))
        {
            throw new InvalidOperationException(errorMessage);
        }
    }

    /// <summary>
    /// Atomically validates a session and marks one operation as active.
    /// </summary>
    public bool TryBeginOperation(
        string sessionId,
        [NotNullWhen(true)] out IExcelBatch? batch,
        [NotNullWhen(false)] out string? errorMessage)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        batch = null;

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            errorMessage = "sessionId is required";
            return false;
        }

        var sessionLock = GetSessionLock(sessionId);
        lock (sessionLock)
        {
            if (_teardownFailures.TryGetValue(sessionId, out var teardownFailure))
            {
                errorMessage = $"Session '{sessionId}' is quarantined after a failed close: " +
                               teardownFailure.SourceException.Message;
                return false;
            }

            if (_closingSessions.ContainsKey(sessionId))
            {
                errorMessage = $"Session '{sessionId}' is closing";
                return false;
            }

            if (!_activeSessions.TryGetValue(sessionId, out batch))
            {
                errorMessage = $"Session '{sessionId}' not found";
                return false;
            }

            if (batch.HasTimedOutOperation)
            {
                errorMessage = $"A previous operation on session '{sessionId}' timed out or was cancelled. " +
                               "Please close the session and reopen the workbook before retrying.";
                batch = null;
                return false;
            }

            if (!batch.IsExcelProcessAlive())
            {
                _logger?.LogWarning("Session {SessionId} has dead Excel process, auto-cleaning up", sessionId);
                CleanupDeadSession(sessionId, batch);
                batch = null;
                errorMessage = $"Excel process for session '{sessionId}' has died. Session has been closed. Please create a new session.";
                return false;
            }

            _activeOperationCounts.AddOrUpdate(sessionId, 1, (_, count) => count + 1);
            errorMessage = null;
            return true;
        }
    }

    /// <summary>
    /// Decrements the active operation count for a session.
    /// Call this when an operation completes (success or failure).
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    public void EndOperation(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return;
        var sessionLock = GetSessionLock(sessionId);
        lock (sessionLock)
        {
            while (_activeOperationCounts.TryGetValue(sessionId, out var count))
            {
                var next = Math.Max(0, count - 1);
                if (_activeOperationCounts.TryUpdate(sessionId, next, count))
                {
                    return;
                }
            }
        }
    }

    /// <summary>
    /// Gets the number of active operations for a session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>Number of active operations, or 0 if session not found</returns>
    public int GetActiveOperationCount(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return 0;
        return _activeOperationCounts.TryGetValue(sessionId, out var count) ? count : 0;
    }

    /// <summary>
    /// Gets whether Excel is visible for a session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if Excel is visible for this session</returns>
    public bool IsExcelVisible(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        return _activeSessions.TryGetValue(sessionId, out var batch) && batch.IsExcelVisible;
    }

    /// <summary>
    /// Updates the visibility flag for a session.
    /// Called by window management commands when Excel visibility changes mid-session.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <param name="visible">New visibility state</param>
    /// <returns>True if session was found and flag updated, false if session not found</returns>
    public bool SetExcelVisible(string sessionId, bool visible)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        if (!_activeSessions.ContainsKey(sessionId)) return false;
        _showExcelFlags[sessionId] = visible;
        return true;
    }

    /// <summary>
    /// Validates whether a session can be closed safely.
    /// Returns information about blocking conditions.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>Validation result with details about any blocking conditions</returns>
    public CloseValidationResult ValidateClose(string sessionId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return new CloseValidationResult(false, false, 0, "Session ID is required");
        }

        if (!_activeSessions.ContainsKey(sessionId))
        {
            return new CloseValidationResult(false, false, 0, $"Session '{sessionId}' not found");
        }

        if (_teardownFailures.TryGetValue(sessionId, out var teardownFailure))
        {
            return new CloseValidationResult(
                true,
                false,
                0,
                $"Session '{sessionId}' is quarantined after a failed close: " +
                teardownFailure.SourceException.Message);
        }

        var activeOps = GetActiveOperationCount(sessionId);
        var isVisible = IsExcelVisible(sessionId);

        if (activeOps > 0)
        {
            return new CloseValidationResult(true, isVisible, activeOps,
                $"Cannot close: {activeOps} operation(s) still running. Wait for operations to complete before closing.");
        }

        return new CloseValidationResult(true, isVisible, 0, null);
    }

    /// <summary>
    /// Closes the specified session with optional save.
    /// If save is true, saves changes before closing to ensure atomic operation.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <param name="save">Whether to save changes before closing (default: false)</param>
    /// <param name="force">Force close even if operations are running (default: false)</param>
    /// <returns>True if session was found and closed, false if session not found</returns>
    /// <exception cref="InvalidOperationException">Save operation failed or operations still running</exception>
    public bool CloseSession(string sessionId, bool save = false, bool force = false)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return false;
        }

        var sessionLock = GetSessionLock(sessionId);
        IExcelBatch batch;
        var resolvedFailedTeardown = false;
        lock (sessionLock)
        {
            if (!_activeSessions.TryGetValue(sessionId, out batch!))
            {
                return false;
            }

            if (_teardownFailures.TryGetValue(sessionId, out var teardownFailure))
            {
                if (!TryConfirmFailedTeardown(batch))
                {
                    teardownFailure.Throw();
                }

                RemoveSessionTracking(sessionId, removeSessionLock: false);
                resolvedFailedTeardown = true;
            }

            if (resolvedFailedTeardown)
            {
                _closingSessions.TryRemove(sessionId, out _);
            }

            // Check for running operations (unless force is true)
            if (!resolvedFailedTeardown && !force)
            {
                var activeOps = GetActiveOperationCount(sessionId);
                if (activeOps > 0)
                {
                    throw new InvalidOperationException(
                        $"Cannot close session '{sessionId}': {activeOps} operation(s) still running. " +
                        "Wait for all operations to complete before closing, or use force=true to close anyway.");
                }
            }

            if (!resolvedFailedTeardown && save && batch.HasTimedOutOperation)
            {
                throw new InvalidOperationException(
                    $"A previous operation on session '{sessionId}' timed out or was cancelled. " +
                    "Close without saving and reopen the workbook before retrying.");
            }

            if (!resolvedFailedTeardown)
            {
                _closingSessions[sessionId] = 0;
            }
        }

        if (resolvedFailedTeardown)
        {
            _sessionLocks.TryRemove(sessionId, out _);
            return true;
        }

        var closeSucceeded = false;
        try
        {
            // Save first if requested (blocks until complete)
            if (save)
            {
                try
                {
                    batch.Save();
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to save session '{sessionId}' before closing: {ex.Message}", ex);
                }
            }

            CloseSessionSync(sessionId, batch);
            closeSucceeded = true;
            return true;
        }
        finally
        {
            if (!closeSucceeded)
            {
                _closingSessions.TryRemove(sessionId, out _);
            }
        }
    }

    private void CloseSessionSync(string sessionId, IExcelBatch batch)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        try
        {
            batch.Dispose();
        }
        catch (Exception ex)
        {
            var closeFailure = new InvalidOperationException(
                $"Failed to dispose session '{sessionId}': {ex.Message}",
                ex);
            _teardownFailures[sessionId] = ExceptionDispatchInfo.Capture(closeFailure);
            throw closeFailure;
        }

        var sessionLock = GetSessionLock(sessionId);
        lock (sessionLock)
        {
            RemoveSessionTracking(sessionId, removeSessionLock: false);
        }

        _sessionLocks.TryRemove(sessionId, out _);
    }

    private static bool TryConfirmFailedTeardown(IExcelBatch batch) =>
        batch is IExcelBatchTeardownState teardownState
        && teardownState.TryConfirmOwnedProcessTeardown();

    /// <summary>
    /// Gets the number of active sessions.
    /// Note: This count may include dead sessions. Use <see cref="GetActiveSessions"/> for accurate count.
    /// </summary>
    public int ActiveSessionCount => _activeSessions.Count;

    /// <summary>
    /// Checks if the Excel process for a session is still alive.
    /// If the session exists but Excel has died, it is automatically cleaned up.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if session exists and Excel process is alive, false otherwise</returns>
    public bool IsSessionAlive(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        var sessionLock = GetSessionLock(sessionId);
        lock (sessionLock)
        {
            if (!_activeSessions.TryGetValue(sessionId, out var batch)) return false;

            if (batch.IsExcelProcessAlive())
            {
                return true;
            }

            // Auto-cleanup dead session
            _logger?.LogWarning("Session {SessionId} has dead Excel process, auto-cleaning up during IsSessionAlive check", sessionId);
            CleanupDeadSession(sessionId, batch);
            return false;
        }
    }

    /// <summary>
    /// Gets all active session IDs.
    /// Note: This property does not filter dead sessions. Use <see cref="GetActiveSessions"/> for filtered results.
    /// </summary>
    public IEnumerable<string> ActiveSessionIds => _activeSessions.Keys.ToList();

    /// <summary>
    /// Returns a snapshot of active sessions with associated workbook paths.
    /// Dead sessions (where Excel process has died) are automatically cleaned up and excluded.
    /// </summary>
    public IReadOnlyList<SessionDescriptor> GetActiveSessions()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var snapshot = new List<SessionDescriptor>(_sessionFilePaths.Count);
        var deadSessions = new List<(string sessionId, IExcelBatch batch)>();

        foreach (var kvp in _sessionFilePaths)
        {
            var sessionId = kvp.Key;

            // Check if session is still alive
            if (_activeSessions.TryGetValue(sessionId, out var batch))
            {
                if (batch.IsExcelProcessAlive())
                {
                    // Get origin and createdAt metadata (defaults for legacy sessions)
                    _sessionOrigins.TryGetValue(sessionId, out var origin);
                    _sessionCreatedAt.TryGetValue(sessionId, out var createdAt);

                    snapshot.Add(new SessionDescriptor(sessionId, kvp.Value, origin, createdAt == default ? null : createdAt));
                }
                else
                {
                    // Mark for cleanup (don't cleanup during iteration)
                    deadSessions.Add((sessionId, batch));
                }
            }
            // If not in _activeSessions but in _sessionFilePaths, skip (orphaned metadata)
        }

        // Clean up dead sessions after iteration
        foreach (var (sessionId, batch) in deadSessions)
        {
            _logger?.LogWarning("Session {SessionId} has dead Excel process, auto-cleaning up during GetActiveSessions", sessionId);
            CleanupDeadSession(sessionId, batch);
        }

        return snapshot;
    }

    /// <summary>
    /// Attempts to get the workbook path associated with a session ID.
    /// </summary>
    public bool TryGetFilePath(string sessionId, [NotNullWhen(true)] out string? filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            filePath = null;
            return false;
        }

        return _sessionFilePaths.TryGetValue(sessionId, out filePath);
    }

    /// <summary>
    /// Atomically reserves a Save As target path for an active session.
    /// </summary>
    /// <param name="sessionId">Active session ID</param>
    /// <param name="filePath">Prospective workbook path</param>
    /// <returns>The normalized reserved path.</returns>
    public string ReserveSessionFilePath(string sessionId, string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var normalizedPath = Path.GetFullPath(filePath);

        lock (_filePathReservationLock)
        {
            if (!_activeSessions.ContainsKey(sessionId))
            {
                throw new KeyNotFoundException($"Session not found: {sessionId}");
            }

            if (_sessionFilePathReservations.ContainsKey(sessionId))
            {
                throw new InvalidOperationException(
                    $"Session '{sessionId}' already has a Save As operation in progress.");
            }

            if (_sessionFilePaths.TryGetValue(sessionId, out var currentPath) &&
                string.Equals(currentPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                _sessionFilePathReservations[sessionId] = normalizedPath;
                return normalizedPath;
            }

            if (!TryClaimFilePath(normalizedPath, sessionId))
            {
                throw new InvalidOperationException(
                    $"File '{normalizedPath}' is already open or reserved by another session.");
            }

            _sessionFilePathReservations[sessionId] = normalizedPath;
        }

        return normalizedPath;
    }

    /// <summary>
    /// Releases a Save As target reservation when the operation did not complete.
    /// </summary>
    /// <param name="sessionId">Active session ID</param>
    /// <param name="filePath">Previously reserved workbook path</param>
    public void ReleaseSessionFilePathReservation(string sessionId, string filePath)
    {
        var normalizedPath = Path.GetFullPath(filePath);
        lock (_filePathReservationLock)
        {
            if (!_sessionFilePathReservations.TryGetValue(sessionId, out var reservedPath) ||
                !string.Equals(reservedPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            _sessionFilePathReservations.TryRemove(sessionId, out _);
            if (_sessionFilePaths.TryGetValue(sessionId, out var currentPath) &&
                string.Equals(currentPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            ReleaseFilePathClaim(normalizedPath, sessionId);
        }
    }

    /// <summary>
    /// Updates the path associated with an active session after Workbook.SaveAs.
    /// The target path must have been reserved before Excel mutates the workbook.
    /// </summary>
    /// <param name="sessionId">Active session ID</param>
    /// <param name="filePath">New workbook path</param>
    public void UpdateSessionFilePath(string sessionId, string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var normalizedPath = Path.GetFullPath(filePath);

        lock (_filePathReservationLock)
        {
            if (!_activeSessions.ContainsKey(sessionId))
            {
                throw new KeyNotFoundException($"Session not found: {sessionId}");
            }

            if (!_sessionFilePathReservations.TryGetValue(sessionId, out var reservedPath) ||
                !string.Equals(reservedPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"File '{normalizedPath}' is not the active Save As reservation for session '{sessionId}'.");
            }

            if (!_activeFilePaths.TryGetValue(normalizedPath, out var ownerSessionId) ||
                !string.Equals(ownerSessionId, sessionId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"File '{normalizedPath}' is not reserved by session '{sessionId}'.");
            }

            if (_sessionFilePaths.TryGetValue(sessionId, out var previousPath) &&
                !string.Equals(previousPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                _activeFilePaths.TryRemove(previousPath, out _);
            }

            _sessionFilePaths[sessionId] = normalizedPath;
        }
    }

    /// <summary>
    /// Disposes all active sessions, auto-saving each one first to prevent data loss.
    /// </summary>
    /// <remarks>
    /// <para><b>CRITICAL:</b> Sessions are auto-saved before disposal to prevent silent data loss
    /// when the service shuts down (e.g., MCP client disconnect, process exit).</para>
    /// <para><b>CRITICAL:</b> Sessions are disposed SEQUENTIALLY to avoid COM threading issues.</para>
    /// <para>Excel COM objects must be disposed on their STA threads. Parallel disposal causes deadlocks.</para>
    /// </remarks>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        // Close all active sessions SEQUENTIALLY to avoid COM threading issues
        // Excel COM objects must be disposed on their STA threads, parallel disposal causes deadlocks
        var sessions = _activeSessions.Values.ToList();
        _activeSessions.Clear();
        _activeFilePaths.Clear();
        _sessionFilePaths.Clear();

        foreach (var session in sessions)
        {
            // Auto-save before disposal to prevent silent data loss.
            // This protects against the common scenario where the MCP client disconnects
            // or the service process exits, which would otherwise discard all unsaved work.
            if (session.IsExcelProcessAlive())
            {
                try
                {
                    using var saveTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                    session.Save(saveTimeout.Token);
                    _logger.LogInformation("Auto-saved session for {Path} before shutdown", session.WorkbookPath);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to auto-save session for {Path} before shutdown (changes may be lost)", session.WorkbookPath);
                }
            }

            try
            {
                // Dispose sequentially - ExcelBatch.Dispose() handles its own Excel cleanup
                // via ExcelShutdownService with proper timeouts and retry logic
                session.Dispose();
            }
            catch (Exception)
            {
                // Best-effort cleanup — continue with remaining sessions
            }
        }
    }
}

/// <summary>
/// Represents a snapshot of an active Excel session managed by <see cref="SessionManager"/>.
/// </summary>
/// <param name="SessionId">Public session identifier shared with clients.</param>
/// <param name="FilePath">Normalized workbook path associated with the session.</param>
/// <param name="Origin">Which client created this session (CLI or MCP).</param>
/// <param name="CreatedAt">When the session was created.</param>
public sealed record SessionDescriptor(
    string SessionId,
    string FilePath,
    SessionOrigin Origin = SessionOrigin.Unknown,
    DateTime? CreatedAt = null);

/// <summary>
/// Indicates which client created a session.
/// </summary>
public enum SessionOrigin
{
    /// <summary>Session origin is unknown (legacy sessions).</summary>
    Unknown = 0,

    /// <summary>Session was created via the CLI.</summary>
    CLI = 1,

    /// <summary>Session was created via the MCP Server.</summary>
    MCP = 2
}

/// <summary>
/// Result of validating whether a session can be closed.
/// </summary>
/// <param name="SessionExists">Whether the session was found.</param>
/// <param name="IsExcelVisible">Whether Excel is visible (show=true).</param>
/// <param name="ActiveOperationCount">Number of operations currently running.</param>
/// <param name="BlockingReason">Reason why close is blocked, or null if close is allowed.</param>
public sealed record CloseValidationResult(
    bool SessionExists,
    bool IsExcelVisible,
    int ActiveOperationCount,
    string? BlockingReason)
{
    /// <summary>
    /// Whether the session can be closed (no blocking conditions).
    /// </summary>
    public bool CanClose =>
        SessionExists
        && ActiveOperationCount == 0
        && BlockingReason == null;
}
