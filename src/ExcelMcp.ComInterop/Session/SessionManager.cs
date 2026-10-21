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
/// <item><b>Between-process operations CAN be parallel:</b> Sessions sharing an Excel process serialize complete callbacks</item>
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
    private static readonly ConcurrentDictionary<ExcelProcessIdentity, byte> _preservedExcelProcesses = new();
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
        _preservedExcelProcesses.TryRemove(identity, out _);
        NotifyTrackedExcelProcessesChanged();
    }

    internal static void PreserveExcelProcess(ExcelProcessIdentity identity) =>
        _preservedExcelProcesses[identity] = 0;

    internal static bool IsExcelProcessPreserved(ExcelProcessIdentity identity) =>
        _preservedExcelProcesses.ContainsKey(identity);

    internal static bool IsExcelProcessTracked(ExcelProcessIdentity identity) =>
        _trackedExcelProcesses.ContainsKey(identity);

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

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Asynchronous process-tracking observers are isolated from each other; failures are traced per subscriber and do not change the already-published process identity.")]
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

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "ProcessExit only reports exact tracked Excel identities; it must not terminate Excel because user workbooks may share a process started by this host.")]
    private static void OnProcessExit(object? sender, EventArgs e)
    {
        ReportExcelProcessesLeftOnHostExit(_trackedExcelProcesses.Keys);
    }

    internal static void ReportExcelProcessesLeftOnHostExitForTests(
        IReadOnlyCollection<ExcelProcessIdentity> identities) =>
        ReportExcelProcessesLeftOnHostExit(identities);

    private static void ReportExcelProcessesLeftOnHostExit(
        IEnumerable<ExcelProcessIdentity> identities)
    {
        foreach (var identity in identities)
        {
            SessionDiagnostics.WriteStdErr(
                $"Excel process {identity.ProcessId} remains open after the host exits; automatic termination is disabled to preserve workbook state.");
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
    private readonly ConcurrentDictionary<string, byte> _readOnlyValidationSessions = new();
    private readonly ConcurrentDictionary<string, bool> _showExcelFlags = new();
    private readonly ConcurrentDictionary<string, SessionOrigin> _sessionOrigins = new();
    private readonly ConcurrentDictionary<string, DateTime> _sessionCreatedAt = new();
    private readonly ConcurrentDictionary<string, string> _sessionFilePathReservations = new();
    private readonly ConcurrentDictionary<string, byte> _pendingSaveAs = new();
    private readonly Dictionary<ExcelProcessIdentity, ExcelBatch> _attachedApplications = new();
    private readonly HashSet<ExcelProcessIdentity> _detachingApplications = [];
    private readonly ConcurrentDictionary<string, (ExcelBatch Owner, OpenWorkbookInfo Workbook)> _openWorkbookSelections = new();
    private readonly object _filePathReservationLock = new();
    private readonly ILogger<SessionManager> _logger;
    private volatile bool _disposed;
    internal Action? BeforeStaleOwnedSessionLockHookForTests { get; set; }

    private object GetSessionLock(string sessionId) =>
        _sessionLocks.GetOrAdd(sessionId, static _ => new object());

    private bool TryClaimFilePath(string normalizedPath, string sessionId)
    {
        lock (_filePathReservationLock)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
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

    private void PublishSession(
        string sessionId,
        string normalizedPath,
        IExcelBatch batch,
        bool show,
        SessionOrigin origin,
        bool readOnlyValidation = false)
    {
        // Publication and shutdown share a short lock; Excel startup and teardown stay outside it.
        lock (_filePathReservationLock)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (!_activeSessions.TryAdd(sessionId, batch))
                throw new InvalidOperationException($"Session ID collision: {sessionId}");
            if (!_sessionFilePaths.TryAdd(sessionId, normalizedPath))
                throw new InvalidOperationException($"Failed to record session metadata for: {sessionId}");
            _activeOperationCounts[sessionId] = 0;
            _showExcelFlags[sessionId] = show;
            _sessionOrigins[sessionId] = origin;
            _sessionCreatedAt[sessionId] = DateTime.UtcNow;
            if (readOnlyValidation)
                _readOnlyValidationSessions[sessionId] = 0;
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

    /// <summary>Discovers already-open workbooks in the current user's Windows session.</summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Discovery returns partial results and the exact per-application inspection errors instead of treating failed reads as missing workbooks.")]
    public async Task<WorkbookDiscoveryResult> DiscoverOpenWorkbooksAsync()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var nativeDiscovery = RunningExcelWindows.Find();
        var workbooks = new ConcurrentBag<OpenWorkbookInfo>();
        var errors = new ConcurrentBag<WorkbookDiscoveryError>(nativeDiscovery.Errors);

        await Task.WhenAll(nativeDiscovery.Windows
            .GroupBy(window => window.Process)
            .Select(group => group.First())
            .Select(window => Task.Run(() =>
            {
                ExcelBatch? application = null;
                try
                {
                    lock (_filePathReservationLock)
                    {
                        ObjectDisposedException.ThrowIf(_disposed, this);
                        if (_detachingApplications.Contains(window.Process))
                        {
                            throw new InvalidOperationException(
                                "The Excel application connection is being released for an owned-session close. Discover again after it completes.");
                        }
                        if (!_attachedApplications.TryGetValue(window.Process, out application!))
                        {
                            application = ExcelBatch.ConnectApplication(window);
                            _attachedApplications.Add(window.Process, application);
                        }
                    }

                    var discovery = application.DiscoverAttachedWorkbooks();
                    foreach (var workbook in discovery.Workbooks)
                    {
                        var existingSessionId = FindExistingWorkbookSession(application, workbook);
                        var selection = workbook with
                        {
                            SessionId = existingSessionId
                        };
                        if (existingSessionId is not null
                            && _activeSessions.TryGetValue(existingSessionId, out var existingSession)
                            && existingSession is AttachedWorkbookBatch attached)
                        {
                            if (UpdateAttachedWorkbookLocation(existingSessionId, attached, workbook) is { } error)
                                errors.Add(error);
                        }
                        workbooks.Add(selection);
                        _openWorkbookSelections[workbook.WorkbookId] = (application, selection);
                    }
                    foreach (var error in discovery.Errors)
                        errors.Add(error);

                    if (discovery.Errors.Count == 0)
                    {
                        var liveIds = discovery.Workbooks
                            .Select(workbook => workbook.WorkbookId)
                            .ToHashSet(StringComparer.Ordinal);
                        lock (_filePathReservationLock)
                        {
                            var staleSessions = _activeSessions
                                .Where(entry => entry.Value is AttachedWorkbookBatch attached
                                    && ReferenceEquals(attached.Owner, application)
                                    && !liveIds.Contains(attached.WorkbookId))
                                .ToArray();
                            foreach (var (sessionId, session) in staleSessions)
                            {
                                if (session is AttachedWorkbookBatch attached)
                                {
                                    attached.Dispose();
                                    RemoveSessionTracking(sessionId, removeSessionLock: true);
                                }
                            }

                            foreach (var selection in _openWorkbookSelections)
                            {
                                if (ReferenceEquals(selection.Value.Owner, application)
                                    && !liveIds.Contains(selection.Key))
                                {
                                    _openWorkbookSelections.TryRemove(selection.Key, out _);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Exception failure = ex;
                    if (application is { HasFailedStartup: true })
                    {
                        try
                        {
                            ReleaseFailedDiscoveryConnection(window.Process, application);
                        }
                        catch (Exception cleanupFailure)
                        {
                            failure = new AggregateException(
                                "Excel application connection and its release both failed.", ex, cleanupFailure);
                        }
                    }
                    errors.Add(new WorkbookDiscoveryError(
                        window.Process.ProcessId,
                        $"Excel application inspection failed: {failure.GetType().Name}: {failure.Message}"));
                }
            })));

        return new WorkbookDiscoveryResult(
            workbooks.OrderBy(workbook => workbook.ProcessId)
                .ThenBy(workbook => workbook.Name, StringComparer.OrdinalIgnoreCase)
                .ToArray(),
            errors.OrderBy(error => error.ProcessId).ToArray());
    }

    private void ReleaseFailedDiscoveryConnection(ExcelProcessIdentity identity, ExcelBatch application)
    {
        lock (_filePathReservationLock)
        {
            if (!_attachedApplications.TryGetValue(identity, out var current)
                || !ReferenceEquals(current, application)
                || !_detachingApplications.Add(identity))
            {
                return;
            }
        }

        try
        {
            application.DetachApplication();
            lock (_filePathReservationLock)
            {
                if (_attachedApplications.TryGetValue(identity, out var current)
                    && ReferenceEquals(current, application))
                {
                    _attachedApplications.Remove(identity);
                }
            }
        }
        finally
        {
            lock (_filePathReservationLock)
                _detachingApplications.Remove(identity);
        }
    }

    private string? FindExistingWorkbookSession(ExcelBatch owner, OpenWorkbookInfo workbook)
    {
        var attachedSession = _activeSessions.FirstOrDefault(entry =>
            entry.Value is AttachedWorkbookBatch attached
            && ReferenceEquals(attached.Owner, owner)
            && string.Equals(attached.WorkbookId, workbook.WorkbookId, StringComparison.Ordinal)).Key;
        if (attachedSession is not null)
            return attachedSession;

        foreach (var (ownedSessionId, candidate) in _activeSessions.ToArray())
        {
            if (candidate is not ExcelBatch owned || owned.ExcelProcessId != workbook.ProcessId)
                continue;

            lock (GetSessionLock(ownedSessionId))
            {
                if (!_activeSessions.TryGetValue(ownedSessionId, out var current)
                    || !ReferenceEquals(current, owned))
                    continue;
                var previousPath = owned.FindOwnedWorkbookPath(workbook);
                if (previousPath is null)
                    continue;

                if (workbook.Location is { } currentPath
                    && (!string.Equals(previousPath, currentPath, StringComparison.OrdinalIgnoreCase)
                        || !_sessionFilePaths.TryGetValue(ownedSessionId, out var trackedPath)
                        || !string.Equals(trackedPath, currentPath, StringComparison.OrdinalIgnoreCase)))
                {
                    var reserved = ReserveSessionFilePath(ownedSessionId, currentPath);
                    try
                    {
                        ExcelBatchWait.Execute(owned, TimeSpan.FromSeconds(2),
                            (_, _) =>
                            {
                                owned.UpdateWorkbookPath(currentPath);
                                return 0;
                            });
                        UpdateSessionFilePath(ownedSessionId, currentPath);
                    }
                    finally
                    {
                        ReleaseSessionFilePathReservation(ownedSessionId, reserved);
                    }
                }
                return ownedSessionId;
            }
        }

        if (workbook.Location is { } location
            && _activeFilePaths.TryGetValue(location, out var sessionId)
            && _activeSessions.TryGetValue(sessionId, out var session)
            && session is not AttachedWorkbookBatch
            && session.ExcelProcessId == workbook.ProcessId)
        {
            if (session is not ExcelBatch ownedSession)
                return sessionId;

            BeforeStaleOwnedSessionLockHookForTests?.Invoke();
            var sessionLock = GetSessionLock(sessionId);
            lock (sessionLock)
            {
                if (_activeSessions.TryGetValue(sessionId, out var currentSession)
                    && ReferenceEquals(currentSession, ownedSession)
                    && GetActiveOperationCount(sessionId) == 0
                    && ownedSession.DetachClosedOwnedWorkbook(location))
                {
                    RemoveSessionTracking(sessionId, removeSessionLock: false);
                    _sessionLocks.TryRemove(sessionId, out _);
                    return null;
                }
            }

            throw new InvalidOperationException(
                $"The workbook at '{location}' was closed and reopened, but its previous session could not be safely detached. " +
                "The workbook remains open; discovery is incomplete for this Excel process.");
        }

        return null;
    }

    private WorkbookDiscoveryError? UpdateAttachedWorkbookLocation(
        string sessionId,
        AttachedWorkbookBatch attached,
        OpenWorkbookInfo workbook)
    {
        var path = workbook.Location ?? workbook.Name;
        lock (_filePathReservationLock)
        {
            if (!_activeSessions.TryGetValue(sessionId, out var session)
                || !ReferenceEquals(session, attached))
            {
                return null;
            }

            var previousPath = _sessionFilePaths.TryGetValue(sessionId, out var trackedPath)
                ? trackedPath
                : attached.FilePath;
            if (previousPath is not null
                && _activeFilePaths.TryGetValue(previousPath, out var previousOwner)
                && string.Equals(previousOwner, sessionId, StringComparison.Ordinal))
            {
                _activeFilePaths.TryRemove(previousPath, out _);
            }

            attached.UpdateDiscoveredWorkbookLocation(workbook.Location, workbook.Name);
            _sessionFilePaths[sessionId] = path;

            if (workbook.Location is null)
                return null;

            if (_activeFilePaths.TryGetValue(workbook.Location, out var currentOwner)
                && !string.Equals(currentOwner, sessionId, StringComparison.Ordinal))
            {
                return new WorkbookDiscoveryError(
                    workbook.ProcessId,
                    $"Workbook '{workbook.Name}' is already reserved by session '{currentOwner}'. " +
                    "The existing attached session remains connected by workbook identity.");
            }

            _activeFilePaths[workbook.Location] = sessionId;
            return null;
        }
    }

    /// <summary>Connects to a previously discovered open workbook without reopening it.</summary>
    public AttachedWorkbookSession AttachWorkbook(string workbookId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookId);
        if (!_openWorkbookSelections.TryGetValue(workbookId, out var selection))
            throw new ArgumentException(
                "Unknown open-workbook selection. Discover open workbooks first.", nameof(workbookId));

        var current = selection.Owner.RefreshAttachedWorkbook(workbookId)
            ?? throw new InvalidOperationException(
                "The selected workbook is no longer confirmed open. Discover open workbooks again.");
        var existingSessionId = FindExistingWorkbookSession(selection.Owner, current);
        var path = current.Location ?? current.Name;
        lock (_filePathReservationLock)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (selection.Owner.ProcessIdentity is { } process
                && _detachingApplications.Contains(process))
            {
                throw new InvalidOperationException(
                    "The Excel application connection is being released. Discover the workbook again after it completes.");
            }
            if (!_openWorkbookSelections.TryGetValue(workbookId, out var liveSelection)
                || !ReferenceEquals(liveSelection.Owner, selection.Owner))
            {
                throw new InvalidOperationException(
                    "The open-workbook selection changed during attachment. Discover open workbooks again.");
            }

            if (existingSessionId is not null && !_activeSessions.ContainsKey(existingSessionId))
            {
                throw new InvalidOperationException(
                    "The existing workbook session changed during attachment. Discover open workbooks again.");
            }
            existingSessionId ??= _activeSessions.FirstOrDefault(entry =>
                entry.Value is AttachedWorkbookBatch attached
                && ReferenceEquals(attached.Owner, selection.Owner)
                && string.Equals(attached.WorkbookId, workbookId, StringComparison.Ordinal)).Key;
            if (existingSessionId is not null)
            {
                return new AttachedWorkbookSession(
                    existingSessionId, current.Name, current.Location, current.AutoSaveOn);
            }

            var sessionId = Guid.NewGuid().ToString("N");
            var claimedPath = current.Location is not null;
            if (claimedPath && !_activeFilePaths.TryAdd(path, sessionId))
            {
                throw new InvalidOperationException(
                    $"Workbook '{path}' is already managed by another session.");
            }

            try
            {
                var batch = new AttachedWorkbookBatch(selection.Owner, current);
                PublishSession(sessionId, path, batch, selection.Owner.IsExcelVisible, SessionOrigin.Unknown);
                return new AttachedWorkbookSession(
                    sessionId, current.Name, current.Location, current.AutoSaveOn);
            }
            catch
            {
                if (claimedPath
                    && _activeFilePaths.TryGetValue(path, out var owner)
                    && string.Equals(owner, sessionId, StringComparison.Ordinal))
                {
                    _activeFilePaths.TryRemove(path, out _);
                }
                throw;
            }
        }
    }

    private static ExcelBatch? GetExecutionOwner(IExcelBatch batch) => batch switch
    {
        ExcelBatch owned => owned,
        AttachedWorkbookBatch attached => attached.Owner,
        _ => null
    };

    /// <summary>Returns the session connection's exact process identity without claiming process ownership.</summary>
    public ExcelProcessIdentity? GetSessionProcessIdentity(string sessionId) =>
        _activeSessions.TryGetValue(sessionId, out var batch)
            ? GetExecutionOwner(batch)?.ProcessIdentity
            : null;

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
        => CreateSessionCore(filePath, show, operationTimeout, operationTimeout, origin);

    /// <summary>
    /// Test-only seam for operation timeout regressions that need a normal Excel startup allowance.
    /// </summary>
    internal string CreateSessionWithTimeouts(
        string filePath,
        bool show,
        TimeSpan? operationTimeout,
        TimeSpan? startupTimeout,
        SessionOrigin origin = SessionOrigin.Unknown)
        => CreateSessionCore(filePath, show, operationTimeout, startupTimeout, origin);

    private string CreateSessionCore(
        string filePath,
        bool show,
        TimeSpan? operationTimeout,
        TimeSpan? startupTimeout,
        SessionOrigin origin)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        string normalizedPath = WorkbookLocation.Normalize(filePath);
        bool isRemote = WorkbookLocation.IsRemote(normalizedPath);
        if (isRemote && !show)
            throw new ArgumentException("SharePoint workbooks require show=true so Excel authentication and permission prompts remain visible.", nameof(show));

        if (!isRemote && !File.Exists(normalizedPath))
        {
            throw new FileNotFoundException($"Excel file not found: {filePath}. To create a new file, use the 'create' action instead of 'open'.", filePath);
        }

        // Generate unique session ID
        string sessionId = Guid.NewGuid().ToString("N");
        if (!TryClaimFilePath(normalizedPath, sessionId))
        {
            throw new InvalidOperationException(
                $"File '{filePath}' is already open in another session or reserved for one. " +
                "Excel cannot open the same file multiple times.");
        }

        ExcelBatch? batch = null;
        try
        {
            // Reject external file-access failures after the in-process path claim,
            // so a duplicate session reports its canonical ownership error instead
            // of being mistaken for an unrelated OS-level file lock.
            if (!isRemote && !FileAccessValidator.IsIrmProtected(normalizedPath))
            {
                FileAccessValidator.ValidateFileNotLocked(normalizedPath);
            }

            // Create batch session using Core API with retry for transient COM failures
            // (e.g., CO_E_SERVER_EXEC_FAILURE when system resources are constrained)
            lock (_filePathReservationLock)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                batch = ExcelSession.BeginOwnedBatch(show, operationTimeout, startupTimeout, normalizedPath);
                PublishSession(sessionId, normalizedPath, batch, show, origin);
            }
            batch.WaitForStartup();

            return sessionId;
        }
        catch (ExcelStartupPendingException ex)
        {
            throw new SessionStartupPendingException(sessionId, ex.Message, ex);
        }
        catch (Exception ex)
        {
            CleanupFailedStartup(sessionId, batch, ex);
            throw new InvalidOperationException($"Failed to create session for '{filePath}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Validates that Excel can open a workbook without creating a public session.
    /// The workbook is opened read-only and closed without saving.
    /// </summary>
    public void ValidateWorkbookOpen(string filePath, TimeSpan? operationTimeout = null)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var normalizedPath = WorkbookLocation.Normalize(filePath);
        var reservationId = Guid.NewGuid().ToString("N");
        if (!TryClaimFilePath(normalizedPath, reservationId))
        {
            throw new InvalidOperationException(
                $"File '{filePath}' is already open in another session or reserved for one.");
        }

        ExcelBatch? batch = null;
        var retained = false;
        try
        {
            if (!WorkbookLocation.IsRemote(normalizedPath))
                FileAccessValidator.ValidateFileNotLocked(normalizedPath);
            lock (_filePathReservationLock)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                batch = ExcelSession.BeginReadOnlyValidation(normalizedPath, operationTimeout, deferStartup: true);
                PublishSession(reservationId, normalizedPath, batch,
                    show: WorkbookLocation.IsRemote(normalizedPath),
                    SessionOrigin.Unknown,
                    readOnlyValidation: true);
            }
            batch.WaitForStartup();
            batch.Close();
        }
        catch (ExcelStartupPendingException ex)
        {
            retained = true;
            throw new SessionStartupPendingException(reservationId, ex.Message, ex);
        }
        catch (Exception ex)
        {
            if (batch is not null && !TryConfirmFailedTeardown(batch))
            {
                retained = true;
                if (batch is not IExcelBatchCloseState { ShutdownCommitted: false })
                    _teardownFailures[reservationId] = ExceptionDispatchInfo.Capture(ex);
                RevealForRecovery(reservationId);
                throw new SessionRetainedException(reservationId,
                    $"Workbook access validation did not complete teardown. Session '{reservationId}' remains owned: {ex.Message}", ex);
            }
            throw;
        }
        finally
        {
            if (!retained)
            {
                RemoveSessionTracking(reservationId, removeSessionLock: true);
                ReleaseFilePathClaim(normalizedPath, reservationId);
            }
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
            lock (_filePathReservationLock)
            {
                ObjectDisposedException.ThrowIf(_disposed, this);
                batch = ExcelBatch.CreateNewWorkbook(normalizedPath, isMacroEnabled,
                    logger: null, show: show, operationTimeout: operationTimeout, deferStartup: true);
                PublishSession(sessionId, normalizedPath, batch, show, origin);
            }
            batch.WaitForStartup();

            return sessionId;
        }
        catch (ExcelStartupPendingException ex)
        {
            throw new SessionStartupPendingException(sessionId, ex.Message, ex);
        }
        catch (Exception ex)
        {
            CleanupFailedStartup(sessionId, batch, ex);
            throw new InvalidOperationException($"Failed to create session for new file '{filePath}': {ex.Message}", ex);
        }
    }

    private void CleanupFailedStartup(string sessionId, ExcelBatch? batch, Exception startupFailure)
    {
        try
        {
            batch?.Dispose();
        }
        catch (Exception teardownFailure)
        {
            var failure = new AggregateException(
                "Excel startup and its teardown both failed.", startupFailure, teardownFailure);
            _teardownFailures[sessionId] = ExceptionDispatchInfo.Capture(failure);
            RevealForRecovery(sessionId);
            throw new SessionRetainedException(sessionId,
                $"Excel startup failed and teardown remains unconfirmed. Session '{sessionId}' remains owned: {startupFailure.Message}",
                failure);
        }
        RemoveSessionTracking(sessionId, removeSessionLock: true);
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "The Excel process is already confirmed dead; disposal errors are logged while removing stale session tracking, not returned as successful workbook operations.")]
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
            _pendingSaveAs.TryRemove(sessionId, out _);
        }

        _activeOperationCounts.TryRemove(sessionId, out _);
        _closingSessions.TryRemove(sessionId, out _);
        _showExcelFlags.TryRemove(sessionId, out _);
        _sessionOrigins.TryRemove(sessionId, out _);
        _sessionCreatedAt.TryRemove(sessionId, out _);
        _readOnlyValidationSessions.TryRemove(sessionId, out _);
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
        [NotNullWhen(false)] out string? errorMessage) =>
        TryBeginOperation(sessionId, out batch, out errorMessage, out _);

    /// <summary>
    /// Atomically validates a session, reports a typed failure state, and marks
    /// one operation as active.
    /// </summary>
    /// <param name="sessionId">Session to validate.</param>
    /// <param name="batch">The validated session batch.</param>
    /// <param name="errorMessage">The validation failure, if any.</param>
    /// <param name="error">The typed validation failure, if any.</param>
    /// <param name="allowTimedOutOperation">
    /// Set only for native Excel show/hide actions, which do not enter Excel's operation queue.
    /// </param>
    public bool TryBeginOperation(
        string sessionId,
        [NotNullWhen(true)] out IExcelBatch? batch,
        [NotNullWhen(false)] out string? errorMessage,
        out SessionOperationError error,
        bool allowTimedOutOperation = false)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        batch = null;
        error = SessionOperationError.None;

        if (string.IsNullOrWhiteSpace(sessionId))
        {
            errorMessage = "sessionId is required";
            error = SessionOperationError.MissingSessionId;
            return false;
        }

        var sessionLock = GetSessionLock(sessionId);
        lock (sessionLock)
        {
            if (_teardownFailures.TryGetValue(sessionId, out var teardownFailure))
            {
                errorMessage = $"Session '{sessionId}' is quarantined after a failed close: " +
                               teardownFailure.SourceException.Message;
                error = SessionOperationError.Quarantined;
                return false;
            }

            if (_closingSessions.ContainsKey(sessionId))
            {
                errorMessage = $"Session '{sessionId}' is closing";
                error = SessionOperationError.Closing;
                return false;
            }

            if (!_activeSessions.TryGetValue(sessionId, out batch))
            {
                errorMessage = $"Session '{sessionId}' not found";
                error = SessionOperationError.NotFound;
                return false;
            }

            if (batch is ExcelBatch { HasFailedStartup: true } failedStartup)
            {
                errorMessage = $"Excel startup for session '{sessionId}' failed: {failedStartup.OperationStatus.ErrorMessage}";
                batch = null;
                error = SessionOperationError.Quarantined;
                return false;
            }

            if (!allowTimedOutOperation && batch.HasTimedOutOperation &&
                GetRefreshState(batch) != WorkbookRefreshState.Ready)
            {
                errorMessage = $"A previous wait on session '{sessionId}' expired. " +
                               "The workbook remains open; execution or readiness verification is still pending.";
                batch = null;
                error = SessionOperationError.TimedOutOrCancelled;
                return false;
            }

            if (!batch.IsExcelProcessAlive())
            {
                _logger?.LogWarning("Session {SessionId} has dead Excel process, auto-cleaning up", sessionId);
                CleanupDeadSession(sessionId, batch);
                batch = null;
                errorMessage = $"Excel process for session '{sessionId}' has died. Session has been closed. Please create a new session.";
                error = SessionOperationError.ExcelProcessDied;
                return false;
            }

            _ = GetOperationStatus(sessionId);
            _activeOperationCounts.AddOrUpdate(sessionId, 1, (_, count) => count + 1);
            errorMessage = null;
            error = SessionOperationError.None;
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

    /// <summary>Records the command and caller's queue cancellation without retaining arguments or data.</summary>
    public void SetOperationContext(string sessionId, string command, CancellationToken callerCancellation = default)
    {
        if (_activeSessions.TryGetValue(sessionId, out var batch) && GetExecutionOwner(batch) is { } owned)
            owned.SetOperationContext(command, callerCancellation);
    }

    /// <summary>Returns bounded execution metadata without calling Excel.</summary>
    public ExcelOperationStatus? GetOperationStatus(string sessionId)
    {
        if (!_activeSessions.TryGetValue(sessionId, out var batch)
            || GetExecutionOwner(batch) is not { } owned)
            return null;
        var status = owned.OperationStatus;
        if (!status.ExecutionPending)
        {
            lock (_filePathReservationLock)
            {
                if (!_disposed && _pendingSaveAs.ContainsKey(sessionId) &&
                    _sessionFilePathReservations.TryGetValue(sessionId, out var reserved))
                {
                    if (string.Equals(batch.WorkbookPath, reserved, StringComparison.OrdinalIgnoreCase))
                        UpdateSessionFilePath(sessionId, reserved);
                    ReleaseSessionFilePathReservation(sessionId, reserved);
                }
            }
        }
        return status;
    }

    /// <summary>Keeps a Save As reservation until the outstanding callback has finished.</summary>
    public void RetainPendingSaveAs(string sessionId) => _pendingSaveAs[sessionId] = 0;

    /// <summary>Reveals only this session's owned Excel window, without a COM wait.</summary>
    public void RevealForRecovery(string sessionId, bool waitExpired = false)
    {
        if (_activeSessions.TryGetValue(sessionId, out var batch) && GetExecutionOwner(batch) is { } owned)
            owned.RevealForRecovery(waitExpired);
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

        if (!_activeSessions.TryGetValue(sessionId, out var batch))
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
        var refreshState = GetRefreshState(batch);
        var isVisible = IsExcelVisible(sessionId);

        if (refreshState == WorkbookRefreshState.DialogOpen)
        {
            return new CloseValidationResult(true, isVisible, activeOps, "dialogOpen")
            {
                RefreshState = refreshState,
                BlockingMessage = ExcelBusyException.DialogGuidance +
                    (batch is IExcelBatchWindowControl { DialogDetails: { } details } ? $" {details}" : "") +
                    (batch is IExcelBatchWindowControl { DialogVisibilityError: { } error } ? $" {error}" : "")
            };
        }

        if (activeOps > 0 && refreshState != WorkbookRefreshState.DialogOpen)
        {
            return new CloseValidationResult(true, isVisible, activeOps,
                $"Cannot close: {activeOps} operation(s) still running. Wait for operations to complete before closing.")
            {
                RefreshState = refreshState == WorkbookRefreshState.Ready ? WorkbookRefreshState.Busy : refreshState
            };
        }

        try
        {
            ExcelBusyException.ThrowIfNotReady(refreshState, "close");
            return new CloseValidationResult(true, isVisible, activeOps, null) { RefreshState = refreshState };
        }
        catch (ExcelBusyException ex)
        {
            return new CloseValidationResult(true, isVisible, activeOps, ex.Message) { RefreshState = refreshState };
        }
    }

    private static WorkbookRefreshState GetRefreshState(IExcelBatch batch) =>
        batch is IExcelBatchRefreshState refreshState
            ? refreshState.GetRefreshState()
            : WorkbookRefreshState.Unknown;

    // Normal close verifies workbook readiness on its STA; recovery still requires the bounded scan.
    private static WorkbookRefreshState GetClosePreflightState(IExcelBatch batch) =>
        batch is ExcelBatch { HasTimedOutOperation: false } owned
            ? owned.GetExecutionPreflightState()
            : GetRefreshState(batch);

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

            if (batch is ExcelBatch { HasFailedStartup: true } && !batch.IsExcelProcessAlive()
                && TryConfirmFailedTeardown(batch))
            {
                RemoveSessionTracking(sessionId, removeSessionLock: false);
                resolvedFailedTeardown = true;
            }
            else if (_teardownFailures.TryGetValue(sessionId, out var teardownFailure))
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

            var refreshState = !resolvedFailedTeardown && !force
                ? GetClosePreflightState(batch)
                : WorkbookRefreshState.Ready;
            if (refreshState == WorkbookRefreshState.DialogOpen)
            {
                ExcelBusyException.ThrowIfNotReady(refreshState, "close");
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

            if (!resolvedFailedTeardown && !force)
            {
                ExcelBusyException.ThrowIfNotReady(refreshState, "close");
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
            ReleaseDiscoveryConnectionBeforeOwnedClose(sessionId, batch);

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

            BeforeCloseSessionHookForTests?.Invoke();
            CloseSessionSync(sessionId, batch, force);
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

    internal Action? BeforeCloseSessionHookForTests { get; set; }

    private void ReleaseDiscoveryConnectionBeforeOwnedClose(string sessionId, IExcelBatch batch)
    {
        if (batch is not ExcelBatch owned || owned.ProcessIdentity is not { } identity)
            return;

        ExcelBatch application;
        AttachedWorkbookBatch[] attachedSessions;
        lock (_filePathReservationLock)
        {
            if (!_attachedApplications.TryGetValue(identity, out application!))
                return;
            if (_detachingApplications.Contains(identity))
            {
                throw new InvalidOperationException(
                    $"Excel process {identity.ProcessId} is already releasing its discovery connection.");
            }

            attachedSessions = _activeSessions
                .Where(entry => entry.Key != sessionId
                    && entry.Value is AttachedWorkbookBatch attached
                    && ReferenceEquals(attached.Owner, application))
                .Select(entry => (AttachedWorkbookBatch)entry.Value)
                .ToArray();
        }

        if (attachedSessions.Length > 0)
        {
            var discovery = application.DiscoverAttachedWorkbooks();
            if (discovery.Errors.Count > 0)
            {
                throw new InvalidOperationException(
                    $"Cannot close owned session '{sessionId}' safely because discovery of Excel process {identity.ProcessId} was incomplete: " +
                    string.Join("; ", discovery.Errors.Select(error => error.Message)));
            }

            var currentWorkbooks = discovery.Workbooks.ToDictionary(
                workbook => workbook.WorkbookId, StringComparer.Ordinal);
            var pathErrors = new List<WorkbookDiscoveryError>();
            foreach (var entry in _activeSessions.ToArray())
            {
                if (entry.Value is not AttachedWorkbookBatch attached
                    || !ReferenceEquals(attached.Owner, application))
                {
                    continue;
                }

                if (currentWorkbooks.TryGetValue(attached.WorkbookId, out var workbook))
                {
                    if (UpdateAttachedWorkbookLocation(entry.Key, attached, workbook) is { } error)
                        pathErrors.Add(error);
                    continue;
                }

                lock (_filePathReservationLock)
                {
                    if (_activeSessions.TryGetValue(entry.Key, out var current)
                        && ReferenceEquals(current, attached))
                    {
                        attached.Dispose();
                        RemoveSessionTracking(entry.Key, removeSessionLock: true);
                    }
                    _openWorkbookSelections.TryRemove(attached.WorkbookId, out _);
                }
            }

            if (pathErrors.Count > 0)
            {
                throw new InvalidOperationException(
                    $"Cannot close owned session '{sessionId}' safely because attached workbook locations conflict with other sessions: " +
                    string.Join("; ", pathErrors.Select(error => error.Message)));
            }
        }

        lock (_filePathReservationLock)
        {
            if (!_attachedApplications.TryGetValue(identity, out var currentApplication)
                || !ReferenceEquals(currentApplication, application))
            {
                return;
            }
            if (_activeSessions.Values.Any(session =>
                    session is AttachedWorkbookBatch attached
                    && ReferenceEquals(attached.Owner, application)))
            {
                return;
            }

            _detachingApplications.Add(identity);
        }

        try
        {
            application.DetachApplication();
        }
        catch
        {
            lock (_filePathReservationLock)
                _detachingApplications.Remove(identity);
            throw;
        }

        lock (_filePathReservationLock)
        {
            if (_attachedApplications.TryGetValue(identity, out var currentApplication)
                && ReferenceEquals(currentApplication, application))
            {
                _attachedApplications.Remove(identity);
            }
            foreach (var selection in _openWorkbookSelections)
            {
                if (ReferenceEquals(selection.Value.Owner, application))
                    _openWorkbookSelections.TryRemove(selection.Key, out _);
            }
            _detachingApplications.Remove(identity);
        }
    }

    private void CloseSessionSync(string sessionId, IExcelBatch batch, bool force)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
        {
            return;
        }

        try
        {
            CloseBatch(batch, force);
        }
        catch (Exception ex) when (ex is ExcelBusyException
            || batch is IExcelBatchCloseState { ShutdownCommitted: false })
        {
            if (batch is ExcelBatch ownedBatch)
                ReleaseClosedWorkbookReservations(sessionId, ownedBatch);
            throw;
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

    private void ReleaseClosedWorkbookReservations(string sessionId, ExcelBatch batch)
    {
        var closedPaths = batch.TakeClosedWorkbookPathsDuringShutdown();
        if (closedPaths.Count == 0)
            return;

        lock (_filePathReservationLock)
        {
            foreach (var path in closedPaths)
            {
                var normalizedPath = WorkbookLocation.Normalize(path);
                if (_activeFilePaths.TryGetValue(normalizedPath, out var owner)
                    && string.Equals(owner, sessionId, StringComparison.Ordinal))
                {
                    _activeFilePaths.TryRemove(normalizedPath, out _);
                }
            }

            if (_sessionFilePaths.ContainsKey(sessionId))
                _sessionFilePaths[sessionId] = WorkbookLocation.Normalize(batch.WorkbookPath);
        }
    }

    private static void CloseBatch(IExcelBatch batch, bool force = false)
    {
        if (batch is AttachedWorkbookBatch attached)
        {
            attached.Close(force);
            return;
        }

        if (!force && !batch.HasTimedOutOperation && batch.IsExcelProcessAlive()
            && batch is IExcelBatchCloseState closeState)
            closeState.Close();
        else
            batch.Dispose();
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
    /// Checks whether a session has live Excel or a retained startup failure.
    /// Other dead sessions are automatically cleaned up.
    /// </summary>
    /// <param name="sessionId">Session ID</param>
    /// <returns>True if the session must remain discoverable, false otherwise</returns>
    public bool IsSessionRetained(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId)) return false;
        var sessionLock = GetSessionLock(sessionId);
        lock (sessionLock)
        {
            if (!_activeSessions.TryGetValue(sessionId, out var batch)) return false;

            if (batch.IsExcelProcessAlive() || batch is ExcelBatch { HasFailedStartup: true })
            {
                _ = GetOperationStatus(sessionId);
                return true;
            }

            // Auto-cleanup dead session
            _logger?.LogWarning("Session {SessionId} has dead Excel process, auto-cleaning up during retention check", sessionId);
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
                if (batch.IsExcelProcessAlive() || batch is ExcelBatch { HasFailedStartup: true })
                {
                    _ = GetOperationStatus(sessionId);
                    // Get origin and createdAt metadata (defaults for legacy sessions)
                    _sessionOrigins.TryGetValue(sessionId, out var origin);
                    _sessionCreatedAt.TryGetValue(sessionId, out var createdAt);

                    var path = _sessionFilePaths.TryGetValue(sessionId, out var currentPath) ? currentPath : kvp.Value;
                    if (batch is AttachedWorkbookBatch attached)
                    {
                        snapshot.Add(new SessionDescriptor(
                            sessionId, attached.FilePath, origin,
                            createdAt == default ? null : createdAt,
                            attached.WorkbookName, IsAttached: true));
                    }
                    else
                    {
                        snapshot.Add(new SessionDescriptor(
                            sessionId, path, origin, createdAt == default ? null : createdAt));
                    }
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
        var normalizedPath = WorkbookLocation.Normalize(filePath);

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
        var normalizedPath = WorkbookLocation.Normalize(filePath);
        lock (_filePathReservationLock)
        {
            if (!_sessionFilePathReservations.TryGetValue(sessionId, out var reservedPath) ||
                !string.Equals(reservedPath, normalizedPath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
            _pendingSaveAs.TryRemove(sessionId, out _);

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
        var normalizedPath = WorkbookLocation.Normalize(filePath);

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
            _readOnlyValidationSessions.TryRemove(sessionId, out _);
        }
    }

    /// <summary>
    /// Disposes all active sessions, auto-saving each one first to prevent data loss.
    /// </summary>
    /// <remarks>
    /// <para>Owned sessions are auto-saved before disposal. Explicitly attached
    /// workbooks are detached without saving or closing their workbook.</para>
    /// <para><b>CRITICAL:</b> Sessions are disposed SEQUENTIALLY to avoid COM threading issues.</para>
    /// <para>Excel COM objects must be disposed on their STA threads. Parallel disposal causes deadlocks.</para>
    /// </remarks>
    /// <exception cref="AggregateException">One or more sessions could not be saved or shut down.</exception>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "All per-session save and teardown failures are collected with their original exception and thrown as AggregateException after attempting the remaining sessions.")]
    public void Dispose()
    {
        KeyValuePair<string, IExcelBatch>[] sessions;
        ExcelBatch[] attachedApplications;
        lock (_filePathReservationLock)
        {
            if (_disposed)
                return;
            _disposed = true;
            sessions = _activeSessions.ToArray();
            attachedApplications = _attachedApplications.Values.ToArray();
        }

        var failures = new List<Exception>();
        foreach (var application in attachedApplications)
        {
            try
            {
                if (!application.QuitOwnedApplicationIfEmpty())
                    application.DetachApplication();
                lock (_filePathReservationLock)
                {
                    var identity = _attachedApplications.FirstOrDefault(
                        entry => ReferenceEquals(entry.Value, application)).Key;
                    if (identity != default)
                        _attachedApplications.Remove(identity);
                    foreach (var selection in _openWorkbookSelections)
                    {
                        if (ReferenceEquals(selection.Value.Owner, application))
                            _openWorkbookSelections.TryRemove(selection.Key, out _);
                    }
                }
            }
            catch (Exception ex)
            {
                failures.Add(new InvalidOperationException(
                    $"Could not release the connection to Excel process {application.ExcelProcessId}; its workbooks were not closed.",
                    ex));
            }
        }

        foreach (var (sessionId, session) in sessions)
        {
            if (session is AttachedWorkbookBatch attached)
            {
                attached.Dispose();
                RemoveSessionTracking(sessionId, removeSessionLock: true);
                continue;
            }

            var teardownFailed = _teardownFailures.ContainsKey(sessionId);
            var readOnlyValidation = _readOnlyValidationSessions.ContainsKey(sessionId);
            if (teardownFailed && TryConfirmFailedTeardown(session))
            {
                RemoveSessionTracking(sessionId, removeSessionLock: true);
                continue;
            }
            // Auto-save before disposal to prevent silent data loss.
            // This protects against the common scenario where the MCP client disconnects
            // or the service process exits, which would otherwise discard all unsaved work.
            if (!teardownFailed && !readOnlyValidation && session.IsExcelProcessAlive())
            {
                try
                {
                    using var saveTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                    session.Save(saveTimeout.Token);
                    _logger.LogInformation("Auto-saved session for {Path} before shutdown", session.WorkbookPath);
                }
                catch (ExcelBusyException ex)
                {
                    RevealForRecovery(sessionId);
                    failures.Add(new InvalidOperationException(
                        $"Shutdown refused for session '{sessionId}': {ex.Message}", ex));
                    _logger.LogWarning(ex,
                        "Shutdown refused for session {SessionId}; workbook and session remain open", sessionId);
                    continue;
                }
                catch (Exception ex)
                {
                    failures.Add(new InvalidOperationException(
                        $"Failed to auto-save session '{sessionId}': {ex.Message}",
                        ex));
                    RevealForRecovery(sessionId);
                    _logger.LogWarning(ex, "Failed to auto-save session for {Path} before shutdown; the workbook remains open", session.WorkbookPath);
                    continue;
                }
            }

            try
            {
                // Dispose sequentially - ExcelBatch.Dispose() handles its own Excel cleanup
                // via ExcelShutdownService with proper timeouts and retry logic
                if (teardownFailed || readOnlyValidation)
                    session.Dispose();
                else
                    CloseBatch(session);
            }
            catch (ExcelBusyException ex)
            {
                RevealForRecovery(sessionId);
                failures.Add(new InvalidOperationException(
                    $"Shutdown refused for session '{sessionId}': {ex.Message}", ex));
                _logger.LogWarning(ex,
                    "Shutdown refused for session {SessionId}; workbook and session remain open", sessionId);
                continue;
            }
            catch (Exception ex)
            {
                var failure = new InvalidOperationException(
                    $"Failed to dispose session '{sessionId}': {ex.Message}", ex);
                failures.Add(failure);
                if (session is not IExcelBatchCloseState { ShutdownCommitted: false })
                    _teardownFailures[sessionId] = ExceptionDispatchInfo.Capture(failure);
                RevealForRecovery(sessionId);
                _logger.LogError(ex, "Failed to dispose session {SessionId}", sessionId);
                continue;
            }
            lock (_filePathReservationLock)
                RemoveSessionTracking(sessionId, removeSessionLock: true);
        }

        lock (_filePathReservationLock)
        {
            if (_activeSessions.IsEmpty && _attachedApplications.Count == 0)
                _activeFilePaths.Clear();
            else
                _disposed = false;
        }

        if (failures.Count > 0)
        {
            throw new AggregateException("One or more Excel sessions failed to shut down.", failures);
        }
    }
}

/// <summary>
/// Represents a snapshot of an active Excel session managed by <see cref="SessionManager"/>.
/// </summary>
/// <param name="SessionId">Public session identifier shared with clients.</param>
/// <param name="FilePath">Workbook location, or null for an attached workbook that has never been saved.</param>
/// <param name="Origin">Which client created this session (CLI or MCP).</param>
/// <param name="CreatedAt">When the session was created.</param>
/// <param name="WorkbookName">Current name for an attached workbook.</param>
/// <param name="IsAttached">Whether the session is attached to a workbook that was already open.</param>
public sealed record SessionDescriptor(
    string SessionId,
    string? FilePath,
    SessionOrigin Origin = SessionOrigin.Unknown,
    DateTime? CreatedAt = null,
    string? WorkbookName = null,
    bool IsAttached = false);

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
    internal WorkbookRefreshState RefreshState { get; init; } = WorkbookRefreshState.Unknown;

    /// <summary>
    /// Actionable explanation for a machine-readable blocking reason, when available.
    /// </summary>
    public string? BlockingMessage { get; init; }

    /// <summary>
    /// Live Excel readiness, including owned modal dialogs; unknown is not proof of readiness.
    /// </summary>
    public string ExcelState => RefreshState switch
    {
        WorkbookRefreshState.Ready => "ready",
        WorkbookRefreshState.Refreshing => "refreshing",
        WorkbookRefreshState.Busy => "busy",
        WorkbookRefreshState.DialogOpen => "dialogOpen",
        _ => "unknown"
    };

    /// <summary>
    /// Whether a server-tracked operation, data refresh, or blocking dialog is still pending.
    /// </summary>
    public bool OperationPending =>
        ActiveOperationCount > 0
        || RefreshState is WorkbookRefreshState.Refreshing or WorkbookRefreshState.DialogOpen;

    /// <summary>
    /// Whether the session can be closed (no blocking conditions).
    /// </summary>
    public bool CanClose =>
        SessionExists
        && ActiveOperationCount == 0
        && BlockingReason == null;
}
