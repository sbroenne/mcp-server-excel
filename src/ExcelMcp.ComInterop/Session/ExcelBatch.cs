using System.Runtime.InteropServices;
using System.Threading.Channels;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Implementation of IExcelBatch that manages a single Excel instance on a dedicated STA thread.
/// Ensures proper COM interop with Excel using STA apartment state and OLE message filter.
/// </summary>
/// <remarks>
/// <para><b>CRITICAL: Excel COM Threading Model</b></para>
/// <list type="bullet">
/// <item>Each ExcelBatch runs on ONE dedicated STA (Single-Threaded Apartment) thread</item>
/// <item>Operations are queued via Channel and executed SERIALLY (never in parallel)</item>
/// <item>Multiple simultaneous Execute() calls are processed one at a time</item>
/// <item>This is a COM interop requirement, not an implementation choice</item>
/// <item>Connections sharing an Excel process serialize complete callbacks, including state restoration</item>
/// </list>
/// <para><b>Resource Cost:</b> Each ExcelBatch = one Excel.Application process (~50-100MB+ memory)</para>
/// </remarks>
internal sealed partial class ExcelBatch : IExcelBatch, IExcelBatchTeardownState, IExcelBatchRefreshState, IExcelBatchCloseState, IExcelBatchWindowControl
{
    // P/Invoke for getting process ID from window handle
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private string _workbookPath; // Primary workbook path
    private long _locationRevision;
    internal long LocationRevision => Interlocked.Read(ref _locationRevision);
    internal void InvalidateLocationSnapshot() => Interlocked.Increment(ref _locationRevision);
    private readonly string[] _allWorkbookPaths; // All workbook paths (includes primary)
    private readonly string[] _openingLocations;
    private readonly bool _showExcel; // Whether to show Excel window
    private readonly bool _openReadOnly; // Whether existing workbooks are opened read-only
    private readonly bool _createNewFile; // Whether to create a new file instead of opening existing
    private readonly bool _isMacroEnabled; // For new files: whether to create .xlsm (macro-enabled)
    private readonly RunningExcelWindows.ApplicationWindow? _attachedApplication;
    private ExcelApplicationLease? _applicationLease;
    private readonly TimeSpan _startupTimeout;
    private readonly TimeSpan _operationTimeout; // Timeout for individual operations
    private readonly ILogger<ExcelBatch> _logger;
    private readonly Channel<IExcelWorkItem> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private readonly Lock _disposeLock = new();
    private int _disposed; // 0 = not disposed, 1 = disposed (using int for Interlocked.CompareExchange)
    private int? _excelProcessId;
    private ExcelProcessIdentity? _excelProcessIdentity;
    private volatile bool _isExcelVisible;
    private volatile bool _operationTimedOut;
    private volatile bool _revealRequested;
    private long? _revealRequestRevision;
    private string? _windowReveal;
    private Task _startupCompletion = Task.CompletedTask;
    private readonly TaskCompletionSource _teardownCompletion = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly object _statusLock = new();
    private readonly Lock _readinessLock = new();
    private bool _teardownConfirmed;
    private bool _staCleanupPending;
    private bool _ownedProcessLeftOpen;
    private bool _quitEmptyOwnedApplication;
    private bool _leaveOwnedApplicationOpen;
    private bool? _originalDisplayAlerts;
    internal Action? BeforeApplicationLeaseReleaseHookForTests { get; set; }
    private readonly List<string> _closedWorkbookPathsDuringShutdown = [];
    private readonly AsyncLocal<(string? Command, CancellationToken CallerCancellation)> _commandContext = new();
    private string? _operationCommand;
    private long _operationStartedAt = System.Diagnostics.Stopwatch.GetTimestamp();
    private DateTime _operationStartedAtUtc = DateTime.UtcNow;
    private double? _completedElapsedSeconds;
    private string _operationState = "starting";
    private string? _operationError;
    private bool _waitExpired;
    private Task<WorkbookRefreshState>? _readinessTask;
    private string? _readinessSelectionId;
    private DateTime _readinessOperationStartedAtUtc;
    private readonly Lock _windowControlLock = new();
    private nint _revealedDialog;
    public string? DialogVisibilityError { get; private set; }
    public string? DialogDetails
    {
        get
        {
            _ = ReadDialogState();
            return ExcelDialogProbe.ReadDetails(_excelProcessIdentity, _logger);
        }
    }

    /// <summary>
    /// Test-only seam for simulating a blocked workbook open path during startup.
    /// Production code must leave this null.
    /// </summary>
    internal static Action<string, CancellationToken>? BeforeWorkbookOpenHook { get; set; }
    internal static Action<object, object>? AfterWorkbookOpenHookForTests { get; set; }
    internal static Func<int, ExcelProcessIdentity?>? TrackProcessIdentityHookForTests { get; set; }

    internal static Func<ExcelProcessIdentity, bool>? FailedStartupExitConfirmationHook { get; set; }

    internal static Action? WorkItemQueuedHookForTests { get; set; }
    internal Action? BeforeRefreshStateReadHookForTests { get; set; }
    internal Action? BeforeVisibilitySnapshotPublicationHookForTests { get; set; }
    internal static Action<ExcelBatch>? BeforeStartupVisibilityPublicationHookForTests { get; set; }
    internal Action? BeforeStaCleanupHookForTests { get; set; }
    internal Action<object>? BeforeEmptyApplicationDetachHookForTests { get; set; }

    // COM state (STA thread only)
    private Excel.Application? _excel;
    private Excel.Workbook? _workbook; // Primary workbook
    private Dictionary<string, Excel.Workbook>? _workbooks; // All workbooks keyed by normalized path
    private ExcelContext? _context;

    private interface IExcelWorkItem
    {
        bool TryExecute();

        bool TryDiscard(Exception? exception = null);
    }

    private sealed class ExcelWorkItem<T>(
        Func<T> operation,
        TaskCompletionSource<T> completion,
        Action<Exception?>? finished = null) : IExcelWorkItem
    {
        private const int Queued = 0;
        private const int Executing = 1;
        private const int Completed = 2;
        private const int Discarded = 3;
        private int _state;

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "STA work-item boundary transfers the exact operation exception to the waiting caller through TaskCompletionSource.TrySetException; throwing on this thread would break failure transport.")]
        public bool TryExecute()
        {
            if (Interlocked.CompareExchange(ref _state, Executing, Queued) != Queued)
            {
                return false;
            }

            try
            {
                var result = operation();
                Volatile.Write(ref _state, Completed);
                finished?.Invoke(null);
                completion.TrySetResult(result);
            }
            catch (Exception ex)
            {
                Volatile.Write(ref _state, Completed);
                finished?.Invoke(ex);
                // Preserve operation-originated cancellation and its inner failures.
                // Execute handles cancellation of the caller's wait separately.
                completion.TrySetException(ex);
            }

            return true;
        }

        public bool TryDiscard(Exception? exception = null)
        {
            if (Interlocked.CompareExchange(ref _state, Discarded, Queued) != Queued)
            {
                return false;
            }

            if (exception != null)
            {
                completion.TrySetException(exception);
            }

            return true;
        }
    }

    /// <summary>
    /// Creates a new ExcelBatch for one or more workbooks.
    /// All workbooks are opened in the same Excel.Application instance, enabling cross-workbook operations.
    /// </summary>
    /// <param name="workbookPaths">Paths to Excel workbooks. First path is the primary workbook.</param>
    /// <param name="logger">Optional logger for diagnostic output. If null, uses NullLogger (no output).</param>
    /// <param name="show">Whether to show the Excel window (default: false for background automation).</param>
    /// <param name="operationTimeout">Timeout for startup and individual operations. Default: 120 seconds.</param>
    /// <param name="openReadOnly">Whether existing workbooks are opened read-only.</param>
    /// <param name="startupTimeout">Internal startup override used by timeout regression tests.</param>
    public ExcelBatch(
        string[] workbookPaths,
        ILogger<ExcelBatch>? logger = null,
        bool show = false,
        TimeSpan? operationTimeout = null,
        bool openReadOnly = false,
        TimeSpan? startupTimeout = null,
        bool deferStartup = false)
        : this(
            workbookPaths,
            logger,
            show,
            openReadOnly,
            createNewFile: false,
            isMacroEnabled: false,
            operationTimeout: operationTimeout,
            startupTimeout: startupTimeout,
            deferStartup: deferStartup)
    {
    }

    /// <summary>
    /// Creates a new ExcelBatch that creates a new workbook file instead of opening an existing one.
    /// The file is saved immediately after creation, then kept open in the session.
    /// </summary>
    /// <param name="filePath">Path where the new Excel file will be created.</param>
    /// <param name="isMacroEnabled">Whether to create .xlsm (macro-enabled) format.</param>
    /// <param name="logger">Optional logger for diagnostic output.</param>
    /// <param name="show">Whether to show the Excel window.</param>
    /// <param name="operationTimeout">Timeout for startup and individual operations. Default: 120 seconds.</param>
    /// <returns>ExcelBatch instance with the new workbook open.</returns>
    internal static ExcelBatch CreateNewWorkbook(string filePath, bool isMacroEnabled, ILogger<ExcelBatch>? logger = null, bool show = false, TimeSpan? operationTimeout = null, bool deferStartup = false)
    {
        return new ExcelBatch(
            [filePath],
            logger,
            show,
            openReadOnly: false,
            createNewFile: true,
            isMacroEnabled: isMacroEnabled,
            operationTimeout: operationTimeout,
            deferStartup: deferStartup);
    }

    internal static ExcelBatch ConnectApplication(RunningExcelWindows.ApplicationWindow window) =>
        new([$"Excel-{window.Process.ProcessId}"], null, false, false, false, false,
            startupTimeout: TimeSpan.FromSeconds(2), deferStartup: true, attachedApplication: window);

    internal AttachedWorkbookDiscovery DiscoverAttachedWorkbooks()
    {
        try
        {
            _startupCompletion.WaitAsync(TimeSpan.FromSeconds(2)).GetAwaiter().GetResult();
        }
        catch (TimeoutException ex)
        {
            throw new TimeoutException(
                $"Connecting to Excel process {_excelProcessId} did not finish within two seconds.", ex);
        }
        _startupCompletion.GetAwaiter().GetResult();
        RecoverTimedOutOperation();
        return ExcelBatchWait.Execute(this, TimeSpan.FromSeconds(2), (_, _) => RefreshAttachedWorkbooks());
    }

    internal OpenWorkbookInfo? RefreshAttachedWorkbook(string workbookId)
    {
        RecoverTimedOutOperation();
        var discovery = ExcelBatchWait.Execute(this, TimeSpan.FromSeconds(2), (_, _) =>
            RefreshAttachedWorkbooks());
        var selection = discovery.Workbooks.FirstOrDefault(workbook =>
            string.Equals(workbook.WorkbookId, workbookId, StringComparison.Ordinal));
        if (selection is not null)
            return selection;
        if (discovery.Errors.Count > 0)
        {
            throw new InvalidOperationException(
                "The selected workbook could not be confirmed because discovery was incomplete: " +
                string.Join("; ", discovery.Errors.Select(error => error.Message)));
        }
        return null;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Workbook COM read failures are returned as explicit per-process discovery errors while successful workbook reads remain available.")]
    private AttachedWorkbookDiscovery RefreshAttachedWorkbooks()
    {
        var retainedWorkbooks = _workbooks
            ?? throw new InvalidOperationException("Attached application workbook ownership is not initialized.");
        var found = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<OpenWorkbookInfo>();
        var errors = new List<WorkbookDiscoveryError>();
        Excel.Workbooks? workbooks = null;
        Excel.ProtectedViewWindows? protectedWindows = null;
        var enumerationComplete = false;

        void RecordWorkbook(ref Excel.Workbook? workbook, bool protectedView)
        {
            var observed = workbook;
            var key = retainedWorkbooks.FirstOrDefault(entry =>
                ReferenceEquals(entry.Value, observed)).Key;
            if (key is null)
            {
                key = Guid.NewGuid().ToString("N");
                retainedWorkbooks.Add(key, workbook!);
                workbook = null;
            }
            if (!found.Add(key))
                return;
            var retained = retainedWorkbooks[key];
            result.Add(new OpenWorkbookInfo(
                key, retained.Name, GetWorkbookLocation(retained.FullName),
                _excelProcessId!.Value, retained.Saved,
                protectedView ? null : ExcelCapabilities.ReadAutoSaveStatus(() => retained.AutoSaveOn),
                ProtectedView: protectedView));
        }

        try
        {
            workbooks = _excel!.Workbooks;
            var count = workbooks.Count;
            enumerationComplete = true;
            for (var index = 1; index <= count; index++)
            {
                Excel.Workbook? workbook = null;
                try
                {
                    workbook = workbooks[index];
                    RecordWorkbook(ref workbook, protectedView: false);
                }
                catch (Exception ex)
                {
                    enumerationComplete = false;
                    errors.Add(new WorkbookDiscoveryError(
                        _excelProcessId,
                        $"Could not inspect workbook {index}: {ex.GetType().Name}: {ex.Message}"));
                }
                finally
                {
                    ComUtilities.Release(ref workbook);
                }
            }

            protectedWindows = _excel.ProtectedViewWindows;
            for (var index = 1; index <= protectedWindows.Count; index++)
            {
                Excel.ProtectedViewWindow? view = null;
                Excel.Workbook? workbook = null;
                try
                {
                    view = protectedWindows[index];
                    workbook = view.Workbook;
                    RecordWorkbook(ref workbook, protectedView: true);
                }
                catch (Exception ex)
                {
                    enumerationComplete = false;
                    errors.Add(new WorkbookDiscoveryError(_excelProcessId,
                        $"Could not inspect Protected View workbook {index}: {ex.GetType().Name}: {ex.Message}"));
                }
                finally
                {
                    ComUtilities.Release(ref workbook);
                    ComUtilities.Release(ref view);
                }
            }

            if (enumerationComplete)
            {
                foreach (var key in retainedWorkbooks.Keys.Except(found).ToArray())
                {
                    var removed = retainedWorkbooks[key];
                    retainedWorkbooks.Remove(key);
                    ComUtilities.Release(ref removed);
                }
            }

            UpdateAttachedPrimaryWorkbook();
        }
        catch (Exception ex)
        {
            errors.Add(new WorkbookDiscoveryError(
                _excelProcessId,
                $"Could not enumerate workbooks: {ex.GetType().Name}: {ex.Message}"));
        }
        finally
        {
            ComUtilities.Release(ref protectedWindows);
            ComUtilities.Release(ref workbooks);
        }

        return new AttachedWorkbookDiscovery(result, errors);
    }

    private static string? GetWorkbookLocation(string fullName)
    {
        return WorkbookLocation.IsRemote(fullName) || Path.IsPathFullyQualified(fullName)
            ? WorkbookLocation.Normalize(fullName)
            : null;
    }

    private void UpdateAttachedPrimaryWorkbook()
    {
        if (_workbooks is null)
            return;
        _workbook = _workbooks.Values.FirstOrDefault();
        _workbookPath = _workbook is null
            ? $"Excel-{_excelProcessId}"
            : GetWorkbookLocation(_workbook.FullName) ?? _workbook.Name;
        _context = _workbook is null
            ? null
            : new ExcelContext(_workbookPath, _excel!, _workbook)
            {
                ProcessIdentity = _excelProcessIdentity
            };
    }

    internal ExcelContext GetAttachedWorkbookContext(string workbookId)
    {
        if (_workbooks is null || !_workbooks.TryGetValue(workbookId, out var workbook))
            throw new InvalidOperationException("The selected workbook is no longer open. Discover open workbooks again.");
        var workbookPath = GetWorkbookLocation(workbook.FullName) ?? workbook.Name;
        return new ExcelContext(workbookPath, _excel!, workbook)
        {
            ProcessIdentity = _excelProcessIdentity
        };
    }

    internal Excel.Workbook GetAttachedWorkbook(string workbookId)
    {
        Excel.Workbook GetSelected()
        {
            if (_workbooks is not null && _workbooks.TryGetValue(workbookId, out var selected))
                return selected;
            throw new InvalidOperationException(
                "The selected workbook is no longer open. Discover open workbooks again.");
        }

        return Thread.CurrentThread == _staThread
            ? GetSelected()
            : Execute((_, _) => GetSelected());
    }

    internal WorkbookRefreshState GetAttachedWorkbookRefreshState(string workbookId) =>
        GetRefreshState(() => ReadAttachedWorkbookRefreshState(workbookId), workbookId);

    private void RecoverTimedOutOperation()
    {
        if (_operationTimedOut && GetRefreshState() != WorkbookRefreshState.Ready)
            throw new ExcelBusyException(
                "A previous Excel operation is still pending or its readiness is unknown. " +
                "The selected workbook remains open; retry discovery after Excel becomes ready.");
    }

    internal void CloseAttachedWorkbook(string workbookId, bool force = false)
    {
        var discovery = RefreshAttachedWorkbooks();
        if (!discovery.Workbooks.Any(workbook =>
                string.Equals(workbook.WorkbookId, workbookId, StringComparison.Ordinal)))
        {
            if (discovery.Errors.Count > 0)
            {
                throw new InvalidOperationException(
                    "The selected workbook could not be confirmed closed because discovery was incomplete: " +
                    string.Join("; ", discovery.Errors.Select(error => error.Message)));
            }
            return;
        }

        var context = GetAttachedWorkbookContext(workbookId);
        if (!force)
            ExcelBusyException.ThrowIfNotReady(ReadSelectedWorkbookRefreshState(context.Book), "close");
        var protectedView = FindProtectedView(context.Book);
        try
        {
            if (protectedView is null)
                ExcelShutdownService.CloseAndQuit(
                    context.Book, null, false, context.WorkbookPath, _logger, releaseWorkbook: false);
            else
                ExcelShutdownService.CloseProtectedViewWindow(protectedView);
        }
        finally
        {
            ComUtilities.Release(ref protectedView);
        }
        var closeDiscovery = RefreshAttachedWorkbooks();
        if (closeDiscovery.Workbooks.Any(workbook =>
                string.Equals(workbook.WorkbookId, workbookId, StringComparison.Ordinal)))
        {
            throw new InvalidOperationException(
                "Excel did not close the selected workbook. It remains open and its session is still attached.");
        }
        if (closeDiscovery.Errors.Count > 0)
        {
            throw new InvalidOperationException(
                "Excel did not confirm that the selected workbook closed. It remains attached to its session: " +
                string.Join("; ", closeDiscovery.Errors.Select(error => error.Message)));
        }

    }

    internal void DetachApplication()
    {
        if (_attachedApplication is null)
            throw new InvalidOperationException("Only an attached application can be detached.");
        lock (_disposeLock)
        {
            if (_teardownConfirmed)
                return;
            lock (_readinessLock)
            {
                lock (_statusLock)
                {
                    if (_disposed == 0)
                    {
                        ThrowIfApplicationWorkPending();
                        Interlocked.Exchange(ref _disposed, 1);
                    }
                }
            }
            CompleteDispose();
        }
    }

    private void ThrowIfApplicationWorkPending()
    {
        if (!HasFailedStartup && (!_startupCompletion.IsCompleted
            || OperationStatus.ExecutionPending || _readinessTask is { IsCompleted: false }))
        {
            throw new ExcelBusyException(
                "Cannot disconnect: Excel work is still pending. The workbook and its session remain connected.");
        }
    }

    internal bool QuitOwnedApplicationIfEmpty()
    {
        if (_disposed != 0 || _attachedApplication is null
            || _excelProcessIdentity is not { } identity
            || !SessionManager.IsExcelProcessTracked(identity)
            || !IsExcelProcessAlive())
        {
            return false;
        }

        ThrowIfApplicationWorkPending();
        RecoverTimedOutOperation();
        var empty = ExcelBatchWait.Execute(this, TimeSpan.FromSeconds(2), (_, _) =>
        {
            Excel.Workbooks? workbooks = null;
            Excel.ProtectedViewWindows? protectedWindows = null;
            try
            {
                workbooks = _excel!.Workbooks;
                protectedWindows = _excel.ProtectedViewWindows;
                return workbooks.Count == 0 && protectedWindows.Count == 0;
            }
            finally
            {
                ComUtilities.Release(ref protectedWindows);
                ComUtilities.Release(ref workbooks);
            }
        });
        if (!empty)
            return false;

        _quitEmptyOwnedApplication = true;
        if (BeforeEmptyApplicationDetachHookForTests is { } beforeDetach)
            Execute((_, _) => beforeDetach(_excel!));
        DetachApplication();
        return true;
    }

    /// <summary>
    /// Private constructor that handles both opening existing files and creating new ones.
    /// </summary>
    private ExcelBatch(
        string[] workbookPaths,
        ILogger<ExcelBatch>? logger,
        bool show,
        bool openReadOnly,
        bool createNewFile,
        bool isMacroEnabled,
        TimeSpan? operationTimeout = null,
        TimeSpan? startupTimeout = null,
        bool deferStartup = false,
        RunningExcelWindows.ApplicationWindow? attachedApplication = null)
    {
        if (workbookPaths == null || workbookPaths.Length == 0)
            throw new ArgumentException("At least one workbook path is required", nameof(workbookPaths));

        _attachedApplication = attachedApplication;

        _openingLocations = workbookPaths;
        _allWorkbookPaths = attachedApplication is null
            ? workbookPaths.Select(WorkbookLocation.Normalize).ToArray()
            : workbookPaths;
        _workbookPath = _allWorkbookPaths[0]; // Primary workbook
        _showExcel = show;
        _openReadOnly = openReadOnly;
        _createNewFile = createNewFile;
        _isMacroEnabled = isMacroEnabled;
        _operationCommand = createNewFile ? "session.create" : openReadOnly ? "session.test" : "session.open";
        _operationTimeout = operationTimeout ?? ComInteropConstants.DefaultOperationTimeout;
        _startupTimeout = startupTimeout ?? _operationTimeout;
        _logger = logger ?? NullLogger<ExcelBatch>.Instance;
        _shutdownCts = new CancellationTokenSource();

        // Create unbounded channel for work items
        _workQueue = Channel.CreateUnbounded<IExcelWorkItem>(new UnboundedChannelOptions
        {
            SingleReader = true,
            SingleWriter = false
        });

        // Start STA thread with message pump
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        _startupCompletion = started.Task;

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "The startup STA reports the original failure through started.TrySetException to the constructor caller. Its final cleanup logs OLE-filter revocation failure without replacing the startup exception.")]
        void RunStaThread()
        {
            Excel.Application? startupExcel = null;
            Excel.Workbook? startupPrimaryWorkbook = null;
            Dictionary<string, Excel.Workbook>? startupWorkbooks = null;
            Exception? executionFailure = null;
            try
            {
                // CRITICAL: Register OLE message filter on STA thread for Excel busy handling
                OleMessageFilter.Register();

                if (_attachedApplication is { } attached)
                {
                    _excelProcessIdentity = attached.Process;
                    _excelProcessId = attached.Process.ProcessId;
                    _applicationLease = new ExcelApplicationLease(attached.Process);
                    using var startupExecution = _applicationLease.EnterExecution();
                    startupExcel = RunningExcelWindows.Connect(attached);
                    _excel = startupExcel;
                    _workbooks = new Dictionary<string, Excel.Workbook>(StringComparer.Ordinal);
                    startupWorkbooks = _workbooks;
                    _operationCommand = "session.discover";
                    var discovery = RefreshAttachedWorkbooks();
                    if (discovery.Errors.Count > 0)
                        _logger.LogWarning(
                            "Initial workbook discovery for attached Excel process {ProcessId} was incomplete: {Errors}",
                            _excelProcessId, string.Join("; ", discovery.Errors.Select(error => error.Message)));
                    var visibilityRevision = _applicationLease.VisibilityRevision;
                    var visible = (bool)startupExcel.Visible;
                    BeforeStartupVisibilityPublicationHookForTests?.Invoke(this);
                    UpdateConfirmedVisibility(visible, visibilityRevision);
                    FinishOperation(null);
                    started.SetResult();
                }
                else
                {
                    // Create Excel and workbook ON THIS STA THREAD
                    Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                    if (excelType == null)
                    {
                        throw new InvalidOperationException("Microsoft Excel is not installed on this system.");
                    }

                    Excel.Application tempExcel;
                    try
                    {
                        tempExcel = ResiliencePipelines.CreateSessionCreationPipeline().Execute(
                            () => (Excel.Application)Activator.CreateInstance(excelType)!);
                    }
                    catch (InvalidCastException ex)
                    {
                        // Enrich with COM environment diagnostics for remote debugging (Issue #559)
                        var diagnostics = ComDiagnostics.FormatForErrorMessage(ComDiagnostics.Collect());
                        throw new InvalidOperationException(
                            $"Failed to cast Excel COM object to PIA interface. " +
                            $"This typically indicates a COM registration mismatch (architecture, version, or corruption).\n{diagnostics}",
                            ex);
                    }
                    startupExcel = tempExcel;
                    // Start Excel visible during workbook open so enterprise auth/sign-in
                    // dialogs are interactable. Hide after all workbooks open successfully.
                    // This also covers IRM-protected files which already needed Visible=true.
                    tempExcel.Visible = true;
                    _originalDisplayAlerts = tempExcel.DisplayAlerts;
                    tempExcel.DisplayAlerts = false;

                    // Readiness probes and teardown require both PID and start time.
                    // Retry when Excel's window or process identity is not available yet.
                    try
                    {
                        const int maxRetries = 3;
                        const int retryDelayMs = 500;

                        for (int attempt = 1; attempt <= maxRetries; attempt++)
                        {
                            int hwnd = tempExcel.Hwnd;
                            if (hwnd != 0)
                            {
                                uint processId = 0;
                                _ = GetWindowThreadProcessId(new IntPtr(hwnd), out processId);
                                if (processId != 0)
                                {
                                    _excelProcessId = (int)processId;
                                    _excelProcessIdentity =
                                        TrackProcessIdentityHookForTests is { } trackIdentity
                                            ? trackIdentity(_excelProcessId.Value)
                                            : SessionManager.TrackExcelProcessIdentity(_excelProcessId.Value);
                                    if (_excelProcessIdentity.HasValue)
                                    {
                                        _logger.LogDebug("Captured Excel process identity via Hwnd: {ProcessId} (attempt {Attempt})",
                                            _excelProcessId, attempt);
                                        break;
                                    }
                                }
                            }

                            if (attempt < maxRetries)
                            {
                                _logger.LogDebug("Excel process identity not available yet (attempt {Attempt}/{Max}), retrying in {Delay}ms",
                                    attempt, maxRetries, retryDelayMs);
                                Thread.Sleep(retryDelayMs);
                            }
                        }

                    }
                    catch (ExcelProcessPersistenceException ex)
                    {
                        _excelProcessIdentity = ex.Identity;
                        throw;
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("Failed to capture Excel process identity during startup.", ex);
                    }
                    if (!_excelProcessIdentity.HasValue)
                    {
                        _logger.LogError("Excel startup cannot continue without a confirmed process identity.");
                        throw new InvalidOperationException(
                            "Could not capture Excel process identity. No workbook has been opened or created, " +
                            "and no session was published. Retry opening the workbook.");
                    }
                    _applicationLease = new ExcelApplicationLease(_excelProcessIdentity.Value);
                    using var startupExecution = _applicationLease.EnterExecution();

                    // Enabling VBA in an already opened workbook requires a separately approved reopen.
                    int automationSecurity = SelectAutomationSecurity(macrosApproved: false);
                    tempExcel.AutomationSecurity = (Microsoft.Office.Core.MsoAutomationSecurity)automationSecurity;

                    // Open or create workbooks in the same Excel instance
                    var tempWorkbooks = new Dictionary<string, Excel.Workbook>(StringComparer.OrdinalIgnoreCase);
                    startupWorkbooks = tempWorkbooks;
                    Excel.Workbook? primaryWorkbook = null;

                    foreach (var path in _openingLocations)
                    {
                        Excel.Workbook wb;
                        string normalizedPath = WorkbookLocation.Normalize(path);

                        if (_createNewFile)
                        {
                            // CREATE NEW FILE: Use Add() + SaveAs() instead of Open()
                            // Validate directory exists (do not create automatically)
                            string? directory = Path.GetDirectoryName(normalizedPath);
                            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                            {
                                throw new DirectoryNotFoundException($"Directory does not exist: '{directory}'. Create the directory first before creating Excel files.");
                            }

                            Excel.Workbooks? workbooks = null;
                            try
                            {
                                workbooks = tempExcel.Workbooks;
                                wb = workbooks.Add();
                            }
                            finally
                            {
                                ComUtilities.Release(ref workbooks);
                            }

                            // SaveAs with appropriate format
                            if (_isMacroEnabled)
                            {
                                wb.SaveAs(normalizedPath, ComInteropConstants.XlOpenXmlWorkbookMacroEnabled);
                            }
                            else
                            {
                                wb.SaveAs(normalizedPath, ComInteropConstants.XlOpenXmlWorkbook);
                            }
                        }
                        else
                        {
                            Excel.Workbooks? workbooks = null;
                            try
                            {
                                BeforeWorkbookOpenHook?.Invoke(path, _shutdownCts.Token);
                                _shutdownCts.Token.ThrowIfCancellationRequested();
                                workbooks = tempExcel.Workbooks;
                                wb = workbooks.Open(
                                    path, UpdateLinks: 0, ReadOnly: _openReadOnly,
                                    IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false);
                            }
                            finally
                            {
                                ComUtilities.Release(ref workbooks);
                            }
                        }

                        tempWorkbooks[normalizedPath] = wb;
                        AfterWorkbookOpenHookForTests?.Invoke(tempExcel, wb);
                        if (normalizedPath == _workbookPath)
                        {
                            primaryWorkbook = wb;
                            startupPrimaryWorkbook = wb;
                        }
                    }

                    // All workbooks opened successfully — safe to apply user's visibility preference.
                    // Enterprise auth/sign-in dialogs (if any) have already been dismissed.
                    tempExcel.Visible = _showExcel || _revealRequested;
                    var visibilityRevision = _applicationLease.VisibilityRevision;
                    UpdateConfirmedVisibility((bool)tempExcel.Visible, visibilityRevision);

                    _excel = tempExcel;
                    _workbook = primaryWorkbook;
                    _workbooks = tempWorkbooks;
                    _context = new ExcelContext(_workbookPath, _excel, _workbook!)
                    {
                        ProcessIdentity = _excelProcessIdentity
                    };

                    FinishOperation(null);
                    started.SetResult();
                }

                // Message pump - process work queue until completion or cancellation.
                // CRITICAL: Uses WaitToReadAsync() instead of polling with Thread.Sleep(10).
                //
                // Why WaitToReadAsync and not polling:
                // 1. Thread.Sleep(10) on an STA thread with registered OLE message filter is unreliable.
                //    Pending COM messages (Excel events during calculation) cause Sleep to return
                //    immediately via MsgWaitForMultipleObjectsEx, turning the loop into a 100% CPU spin.
                // 2. The previous outer catch(Exception){} silently bypassed Thread.Sleep when any
                //    exception occurred, causing tight spin loops with zero backoff.
                // 3. WaitToReadAsync().AsTask().GetAwaiter().GetResult() blocks the thread efficiently
                //    and wakes instantly when work arrives. No COM message pumping occurs during the
                //    block, but that's fine — we don't host COM objects or subscribe to Excel events,
                //    so no inbound COM messages need dispatching while idle. COM calls within work items
                //    pump messages internally via CoWaitForMultipleHandles.
                while (true)
                {
                    try
                    {
                        // Block until work is available, channel completes, or shutdown is requested.
                        if (!_workQueue.Reader.WaitToReadAsync(_shutdownCts.Token)
                                              .AsTask().GetAwaiter().GetResult())
                        {
                            // Channel completed (writer called Complete()) — exit gracefully
                            _logger.LogDebug("Channel completed, exiting message pump for {FileName}", Path.GetFileName(_workbookPath));
                            break;
                        }

                        // Drain all available work items before blocking again
                        while (_workQueue.Reader.TryRead(out var work))
                        {
                            if (_shutdownCts.IsCancellationRequested)
                            {
                                work.TryDiscard(new ObjectDisposedException(
                                    nameof(ExcelBatch),
                                    $"Session for '{Path.GetFileName(_workbookPath)}' was disposed before the queued operation started."));
                                continue;
                            }

                            using var execution = _applicationLease!.EnterExecution();
                            work.TryExecute();
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // Shutdown requested via _shutdownCts.
                        // Complete queued callers without invoking callbacks during shutdown.
                        while (_workQueue.Reader.TryRead(out var remainingWork))
                        {
                            remainingWork.TryDiscard(new ObjectDisposedException(
                                nameof(ExcelBatch),
                                $"Session for '{Path.GetFileName(_workbookPath)}' was disposed before the queued operation started."));
                        }

                        _logger.LogDebug("Shutdown requested, exiting message pump for {FileName}", Path.GetFileName(_workbookPath));
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                executionFailure = ex;
                lock (_statusLock)
                    _staCleanupPending = true;
                FinishOperation(ex);
                started.TrySetException(ex);
            }
            finally
            {
                // Cleanup COM objects on STA thread exit
                _logger.LogDebug("STA thread cleanup starting for {FileName}", Path.GetFileName(_workbookPath));

                // Startup failures can happen after Excel was created but before the local COM state
                // was promoted into the instance fields (for example: locked workbook, Open() failure,
                // or second workbook failure in a multi-workbook batch). In that case, the instance
                // fields remain null and cleanup must fall back to the startup locals to avoid leaking
                // a hidden Excel.exe process.
                try
                {
                    BeforeStaCleanupHookForTests?.Invoke();
                    using var cleanupExecution = _applicationLease?.EnterExecution();
                    var cleanupExcel = _excel ?? startupExcel;
                    var cleanupWorkbooks = _workbooks ?? startupWorkbooks;
                    var cleanupWorkbook = _workbook
                        ?? cleanupWorkbooks?.Values.FirstOrDefault()
                        ?? (_workbooks is null ? startupPrimaryWorkbook : null);

                    // INSTRUMENTATION: Check if Excel process is alive BEFORE entering shutdown
                    if (_excelProcessId.HasValue)
                    {
                        try
                        {
                            using var beforeProc = System.Diagnostics.Process.GetProcessById(_excelProcessId.Value);
                            bool beforeAlive = !beforeProc.HasExited;
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-SHUTDOWN-ENTER-PROCESS-CHECK] Excel PID {_excelProcessId.Value} alive={beforeAlive} file={Path.GetFileName(_workbookPath)}");
                            _logger.LogDebug(
                                "[DIAG-SHUTDOWN-ENTER-PROCESS-CHECK] Excel PID {ProcessId} alive={Alive} file={FileName}",
                                _excelProcessId.Value, beforeAlive, Path.GetFileName(_workbookPath));
                        }
                        catch (ArgumentException)
                        {
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-SHUTDOWN-ENTER-PROCESS-DEAD] Excel PID {_excelProcessId.Value} already dead BEFORE shutdown file={Path.GetFileName(_workbookPath)}");
                            _logger.LogWarning(
                                "[DIAG-SHUTDOWN-ENTER-PROCESS-DEAD] Excel PID {ProcessId} already dead BEFORE shutdown for {FileName}",
                                _excelProcessId.Value, Path.GetFileName(_workbookPath));
                        }
                        catch (System.ComponentModel.Win32Exception ex)
                        {
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-SHUTDOWN-ENTER-PROCESS-INACCESSIBLE] Excel PID {_excelProcessId.Value} could not be queried before shutdown: {ex.Message}");
                            _logger.LogWarning(
                                ex,
                                "[DIAG-SHUTDOWN-ENTER-PROCESS-INACCESSIBLE] Excel PID {ProcessId} could not be queried before shutdown for {FileName}",
                                _excelProcessId.Value,
                                Path.GetFileName(_workbookPath));
                        }
                        catch (InvalidOperationException ex)
                        {
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-SHUTDOWN-ENTER-PROCESS-INACCESSIBLE] Excel PID {_excelProcessId.Value} state was unavailable before shutdown: {ex.Message}");
                            _logger.LogWarning(
                                ex,
                                "[DIAG-SHUTDOWN-ENTER-PROCESS-INACCESSIBLE] Excel PID {ProcessId} state was unavailable before shutdown for {FileName}",
                                _excelProcessId.Value,
                                Path.GetFileName(_workbookPath));
                        }
                    }

                    if (_attachedApplication is null && !_leaveOwnedApplicationOpen)
                    {
                        // Owned applications are closed through the normal shutdown policy.
                        bool quitConfirmed;
                        if (cleanupWorkbooks != null && cleanupWorkbooks.Count > 1)
                        {
                            _logger.LogDebug("Closing {Count} workbooks via ExcelShutdownService", cleanupWorkbooks.Count);

                            foreach (var kvp in cleanupWorkbooks.ToList())
                            {
                                if (kvp.Value == cleanupWorkbook)
                                    continue;
                                ExcelShutdownService.CloseAndQuit(kvp.Value, null, false, kvp.Key, _logger);
                            }
                            cleanupWorkbooks.Clear();
                            quitConfirmed = ExcelShutdownService.CloseAndQuitAndReport(
                                cleanupWorkbook, cleanupExcel, false, _workbookPath, _logger,
                                originalDisplayAlerts: _originalDisplayAlerts);
                        }
                        else
                        {
                            quitConfirmed = ExcelShutdownService.CloseAndQuitAndReport(
                                cleanupWorkbook, cleanupExcel, false, _workbookPath, _logger,
                                originalDisplayAlerts: _originalDisplayAlerts);
                        }
                        _ownedProcessLeftOpen = cleanupExcel is not null && !quitConfirmed;
                    }
                    else
                    {
                        if (cleanupWorkbooks is not null)
                        {
                            foreach (var workbook in cleanupWorkbooks.Values)
                            {
                                var retained = workbook;
                                ComUtilities.Release(ref retained);
                            }
                            cleanupWorkbooks.Clear();
                        }
                        if (_attachedApplication is not null && _quitEmptyOwnedApplication)
                        {
                            _ownedProcessLeftOpen = cleanupExcel is not null
                                && !ExcelShutdownService.CloseAndQuitAndReport(
                                    null, cleanupExcel, false, _workbookPath, _logger);
                        }
                        else
                        {
                            if (cleanupExcel is not null && _originalDisplayAlerts is { } displayAlerts)
                                cleanupExcel.DisplayAlerts = displayAlerts;
                            ComUtilities.Release(ref cleanupExcel);
                        }
                    }

                    _workbook = null;
                    _excel = null;
                    _workbooks = null;
                    _context = null;
                    cleanupWorkbooks?.Clear();
                    cleanupWorkbook = null;
                    cleanupExcel = null;
                    startupWorkbooks?.Clear();
                    startupWorkbooks = null;
                    startupPrimaryWorkbook = null;
                    startupExcel = null;
                    Marshal.CleanupUnusedObjectsInCurrentContext();
                    BeforeApplicationLeaseReleaseHookForTests?.Invoke();
                    _applicationLease?.Dispose();
                    _applicationLease = null;
                    _teardownCompletion.TrySetResult();
                    if (_disposed != 0)
                        FinishOperation(executionFailure);
                }
                catch (Exception teardownFailure)
                {
                    RevealForRecovery();
                    FinishOperation(executionFailure is null
                        ? teardownFailure
                        : new AggregateException("Excel execution and its cleanup both failed.",
                            executionFailure, teardownFailure));
                    _teardownCompletion.TrySetException(teardownFailure);
                    _logger.LogError(teardownFailure,
                        "Excel close/quit failed; its process remains owned and may be partially closed");
                }

                try
                {
                    OleMessageFilter.Revoke();
                }
                catch (Exception ex)
                {
                    // Guard against P/Invoke failure in finally — don't suppress original exception
                    _logger.LogWarning(ex, "OleMessageFilter.Revoke() failed during STA cleanup");
                }

                _logger.LogDebug("STA thread cleanup completed for {FileName}", Path.GetFileName(_workbookPath));
                lock (_statusLock)
                    _staCleanupPending = false;
            }
        }

        _staThread = new Thread(RunStaThread)
        {
            IsBackground = true,
            Name = $"ExcelBatch-{Path.GetFileName(_workbookPath)}"
        };

        // CRITICAL: Set STA apartment state before starting thread
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();
        if (!deferStartup)
            WaitForStartup();
    }

    internal void WaitForStartup()
    {
        // Wait for STA thread to initialize. If startup fails, the STA thread may still be in
        // its finally-block cleanup path (closing a hidden Excel instance created before the
        // failure). Wait for that cleanup before rethrowing so callers do not immediately race
        // a lingering hidden Excel.exe from the failed startup attempt.
        try
        {
            bool completedInTime;
            try
            {
                completedInTime = _startupCompletion.Wait(_startupTimeout);
            }
            catch (AggregateException)
            {
                // Task.Wait() wraps faulted-task exceptions in AggregateException.
                // Fall through to GetAwaiter().GetResult() which unwraps to the
                // original exception type (e.g. InvalidOperationException for IRM).
                completedInTime = true;
            }

            if (!completedInTime)
            {
                _operationTimedOut = true;
                RevealForRecovery(waitExpired: true);
                throw new ExcelStartupPendingException(CreateStartupTimeoutMessage());
            }

            _startupCompletion.GetAwaiter().GetResult();
        }
        catch (Exception startupFailure) when (startupFailure is not ExcelStartupPendingException)
        {
            if (_staThread.IsAlive)
            {
                _ = _staThread.Join(TimeSpan.FromSeconds(10));
            }

            if (_attachedApplication is null && _excelProcessIdentity is { } failedStartupIdentity)
            {
                try
                {
                    FinalizeOwnedProcessTeardown(
                        failedStartupIdentity,
                        FailedStartupExitConfirmationHook
                        ?? WaitForOwnedProcessExit);
                }
                catch (Exception teardownFailure)
                {
                    RevealForRecovery();
                    throw new AggregateException(
                        $"Excel startup failed and exact process teardown could not be confirmed for " +
                        $"process {failedStartupIdentity.ProcessId}. The process identity remains tracked " +
                        "and has not been terminated.",
                        startupFailure, teardownFailure);
                }

                if (_staThread.IsAlive)
                {
                    _ = _staThread.Join(TimeSpan.FromSeconds(5));
                }
            }

            throw;
        }
    }

    private string CreateStartupTimeoutMessage()
    {
        return
            $"Excel startup timed out after {_startupTimeout.TotalSeconds} seconds while opening '{Path.GetFileName(_workbookPath)}'. " +
            "The workbook may be blocked on an interactive dialog, enterprise authentication, IRM/AIP prompt, external-link prompt, or an unresponsive open. " +
            $"The original startup remains owned and may still complete. Do not repeat the open/create to determine its outcome. " +
            $"Excel window: {_windowReveal}. Inspect session status for the eventual result.";
    }

    internal static int SelectAutomationSecurity(bool macrosApproved) => macrosApproved ? 1 : 3;

    public string WorkbookPath => _workbookPath;

    public void UpdateWorkbookPath(string workbookPath) => UpdateOwnedWorkbookPath(_workbookPath, workbookPath);

    internal void UpdateOwnedWorkbookPath(string previousWorkbookPath, string workbookPath)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
        var previousPath = WorkbookLocation.Normalize(previousWorkbookPath);
        var normalizedPath = WorkbookLocation.Normalize(workbookPath);

        void UpdatePath()
        {
            if (_workbooks == null || _excel == null || _workbook == null)
            {
                throw new InvalidOperationException("Workbooks not initialized");
            }

            var index = Array.FindIndex(_allWorkbookPaths,
                path => string.Equals(path, previousPath, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                throw new InvalidOperationException($"Tracked workbook path '{previousPath}' was not found.");
            if (!_workbooks.Remove(previousPath, out var workbook))
            {
                throw new InvalidOperationException($"Tracked workbook '{previousPath}' was not found.");
            }

            _workbooks[normalizedPath] = workbook;
            _allWorkbookPaths[index] = normalizedPath;
            if (index == 0)
            {
                _workbookPath = normalizedPath;
                _context = new ExcelContext(normalizedPath, _excel, _workbook, _context!.Capabilities)
                {
                    ProcessIdentity = _excelProcessIdentity
                };
            }
            InvalidateLocationSnapshot();
        }

        if (Thread.CurrentThread == _staThread)
            UpdatePath();
        else
            Execute((_, _) => UpdatePath());
    }

    public ILogger Logger => _logger;

    public int? ExcelProcessId => _excelProcessId;

    internal ExcelProcessIdentity? ProcessIdentity => _excelProcessIdentity;

    public TimeSpan OperationTimeout => _operationTimeout;

    public bool IsExcelVisible => _applicationLease?.IsExcelVisible ?? _isExcelVisible;

    private void UpdateConfirmedVisibility(bool visible, long? expectedRevision = null)
    {
        _isExcelVisible = visible;
        if (_applicationLease is { } lease)
            lease.UpdateVisibility(visible, expectedRevision);
    }

    public bool HasTimedOutOperation => _operationTimedOut;

    bool IExcelBatchCloseState.ShutdownCommitted => Volatile.Read(ref _disposed) != 0;

    internal bool HasFailedStartup => _startupCompletion.IsFaulted;

    internal ExcelOperationStatus OperationStatus
    {
        get
        {
            lock (_statusLock)
                return new ExcelOperationStatus(
                    _operationState,
                    _completedElapsedSeconds ?? System.Diagnostics.Stopwatch.GetElapsedTime(_operationStartedAt).TotalSeconds,
                    _waitExpired,
                    _staCleanupPending || _operationState is "starting" or "running" or "closing",
                    _operationError,
                    _windowReveal)
                {
                    Command = _operationCommand,
                    StartedAtUtc = _operationStartedAtUtc
                };
        }
    }

    internal void RevealForRecovery(bool waitExpired = false)
    {
        lock (_statusLock)
        {
            if (waitExpired)
            {
                _waitExpired = true;
                if (_operationState is "starting" or "running" or "closing")
                    _operationTimedOut = true;
            }
        }
        string outcome;
        lock (_applicationLease?.WindowControlLock ?? _windowControlLock)
        {
            outcome = ExcelDialogProbe.Reveal(_excelProcessIdentity, _logger);
            if (outcome is "shown" or "alreadyVisible")
                UpdateConfirmedVisibility(true);
            _revealRequestRevision = _applicationLease?.VisibilityRevision;
            _revealRequested = true;
        }
        lock (_statusLock)
            _windowReveal = outcome;
    }

    internal void SetOperationContext(string command, CancellationToken callerCancellation) =>
        _commandContext.Value = (command, callerCancellation);

    private void FinishOperation(Exception? error)
    {
        lock (_statusLock)
        {
            _operationState = error is null ? "completed" : "failed";
            _operationError = error?.Message;
            _completedElapsedSeconds = System.Diagnostics.Stopwatch.GetElapsedTime(_operationStartedAt).TotalSeconds;
        }
    }

    public bool IsExcelProcessAlive()
    {
        if (!_excelProcessId.HasValue)
        {
            // PID capture is best-effort during startup. Under load, Excel can still be healthy
            // even when Hwnd-based PID discovery misses the short retry window. Treat this as
            // "unknown but assumed alive" so callers do not tear down a live session before the
            // first COM command runs; real COM failures still surface on the actual operation.
            return _disposed == 0 || _staThread.IsAlive;
        }

        if (_excelProcessIdentity is { } identity)
        {
            return OwnedProcessGuard.IsAlive(identity);
        }

        try
        {
            using var proc = System.Diagnostics.Process.GetProcessById(_excelProcessId.Value);
            return !proc.HasExited;
        }
        catch (ArgumentException)
        {
            // Process ID doesn't exist - process has terminated
            return false;
        }
    }

    public IReadOnlyDictionary<string, Excel.Workbook> Workbooks
    {
        get
        {
            ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
            return _workbooks ?? throw new InvalidOperationException("Workbooks not initialized");
        }
    }

    public Excel.Workbook GetWorkbook(string filePath)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));

        if (_workbooks == null)
            throw new InvalidOperationException("Workbooks not initialized");

        string normalizedPath = WorkbookLocation.Normalize(filePath);
        if (_workbooks.TryGetValue(normalizedPath, out var workbook))
        {
            return workbook;
        }

        throw new KeyNotFoundException($"Workbook '{filePath}' is not open in this batch.");
    }

    /// <summary>
    /// Executes a void COM operation on the STA thread.
    /// Use this overload for operations that don't need to return values.
    /// All Excel COM operations are synchronous.
    /// </summary>
    public void Execute(
        Action<ExcelContext, CancellationToken> operation,
        CancellationToken cancellationToken = default)
    {
        // Delegate to generic Execute<T> with dummy return
        Execute((ctx, ct) =>
        {
            operation(ctx, ct);
            return 0;
        }, cancellationToken);
    }

    /// <summary>
    /// Executes a COM operation on the STA thread.
    /// All Excel COM operations are synchronous.
    /// </summary>
    public T Execute<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
        cancellationToken.ThrowIfCancellationRequested();
        var (command, callerCancellation) = _commandContext.Value;
        callerCancellation.ThrowIfCancellationRequested();

        // Fail fast if a previous operation timed out or was cancelled while the STA thread
        // was stuck in IDispatch.Invoke. The STA thread cannot process new work items until
        // the hung COM call returns (which may be never). Without this check, new callers
        // would queue work and block until their own timeout expires.
        if (_operationTimedOut)
        {
            throw new TimeoutException(
                $"A previous wait expired for '{Path.GetFileName(_workbookPath)}'. " +
                "The workbook remains open; the original operation or readiness verification is still pending.");
        }
        if (!_startupCompletion.IsCompleted)
            throw new ExcelBusyException("Excel startup is still pending. The workbook remains owned.");
        _startupCompletion.GetAwaiter().GetResult();

        // Check if Excel process is still alive before attempting operation
        if (!IsExcelProcessAlive())
        {
            _logger.LogError("Excel process is no longer running for workbook {FileName}", Path.GetFileName(_workbookPath));
            throw new InvalidOperationException(
                $"Excel process is no longer running for workbook '{Path.GetFileName(_workbookPath)}'. " +
                "The Excel application may have been closed manually or crashed. " +
                "Please close this session and create a new one.");
        }

        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var started = false;
        var workItem = new ExcelWorkItem<T>(() =>
        {
            if (_operationTimedOut)
                throw new ExcelBusyException("An earlier operation exceeded its wait. This queued operation was not executed.");
            lock (_statusLock)
            {
                ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
                started = true;
                _operationStartedAt = System.Diagnostics.Stopwatch.GetTimestamp();
                _operationStartedAtUtc = DateTime.UtcNow;
                _completedElapsedSeconds = null;
                _operationState = "running";
                _operationCommand = command;
                _operationError = null;
                _waitExpired = false;
                _windowReveal = null;
            }

            var application = _excel
                ?? throw new InvalidOperationException("Excel application not initialized.");
            using var writeGuard = new ExcelWriteGuard(application, _logger);

            Exception? operationFailure = null;
            try
            {
                return operation(_context!, CancellationToken.None);
            }
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                UpdateVisibilitySnapshot(operationFailure);
            }
        }, completion, error =>
        {
            if (started)
                FinishOperation(error);
        });

        using var queuedCancellation = callerCancellation.Register(() =>
        {
            if (workItem.TryDiscard(new OperationCanceledException(callerCancellation)))
                RevealForRecovery();
        });
        if (!_workQueue.Writer.TryWrite(workItem))
        {
            throw new ObjectDisposedException(nameof(ExcelBatch),
                $"Session for '{Path.GetFileName(_workbookPath)}' was disposed while submitting an operation.");
        }
        WorkItemQueuedHookForTests?.Invoke();

        // Wait for operation to complete with timeout.
        // When the caller provides a cancellation token (e.g., PowerQuery refresh with its own timeout),
        // respect it exclusively and don't layer the session _operationTimeout on top.
        // This prevents a double-cap where min(callerTimeout, sessionTimeout) is always the shorter one —
        // which caused heavy Power Query refreshes (~8+ min) to always fail against the session default.
        using var timeoutCts = cancellationToken.CanBeCanceled
            ? null
            : new CancellationTokenSource(_operationTimeout);
        try
        {
            return WaitWithDialogMonitoringAsync(completion.Task, timeoutCts?.Token ?? cancellationToken).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException ex) when (
            timeoutCts is { IsCancellationRequested: true } && ex.CancellationToken == timeoutCts.Token)
        {
            if (completion.Task.IsCompleted)
                return completion.Task.GetAwaiter().GetResult();
            // Only cancellation of this wait's actual deadline is a session timeout.
            var expiredWhileQueued = workItem.TryDiscard();
            if (!expiredWhileQueued && !completion.Task.IsCompleted)
            {
                _operationTimedOut = true;
                RevealForRecovery(waitExpired: true);
            }

            if (expiredWhileQueued)
            {
                if (completion.Task.IsCompleted)
                    return completion.Task.GetAwaiter().GetResult();
                _logger.LogError(
                    "Queued operation expired after {Timeout} before execution for {FileName}",
                    _operationTimeout,
                    Path.GetFileName(_workbookPath));
            }
            else
            {
                _logger.LogError(
                    "Operation timed out after {Timeout} for {FileName}",
                    _operationTimeout,
                    Path.GetFileName(_workbookPath));
            }
            throw CreateOperationTimeoutException(expiredWhileQueued);
        }
        catch (OperationCanceledException)
        {
            var cancelledWhileQueued = workItem.TryDiscard();
            _logger.LogDebug("Operation cancelled or timed out for {FileName}", Path.GetFileName(_workbookPath));
            if (!cancelledWhileQueued && !completion.Task.IsCompleted)
            {
                _operationTimedOut = true;
                RevealForRecovery(waitExpired: true);
            }

            throw;
        }
    }

    internal TimeoutException CreateOperationTimeoutException(bool expiredWhileQueued)
    {
        if (expiredWhileQueued)
        {
            return new TimeoutException(
                $"Excel operation expired in the session queue after {_operationTimeout.TotalSeconds} seconds for '{Path.GetFileName(_workbookPath)}' and was not executed.");
        }

        // Inspect native windows directly; the Excel thread may be blocked by the prompt.
        var dialogState = ExcelDialogProbe.Read(
            _excelProcessIdentity, _logger, includeEnabledOwnerOfficeDialog: true);
        var dialogOpen = dialogState == WorkbookRefreshState.DialogOpen;
        return new TimeoutException(
            $"Excel operation timed out after {_operationTimeout.TotalSeconds} seconds for '{Path.GetFileName(_workbookPath)}'. " +
            $"The workbook remains open and the operation may still change it. No changes were saved or discarded by timeout recovery. Excel window: {_windowReveal}. " +
            (dialogOpen
                ? "Excel had a dialog open when the operation timed out and may have been waiting for input. " +
                  "This timeout does not itself cancel the operation or undo workbook changes. " +
                  ExcelDialogProbe.ReadDetails(
                      _excelProcessIdentity, _logger, includeEnabledOwnerOfficeDialog: true)
                : "Excel may be unresponsive or the operation is taking longer than expected. " +
                  "Inspect session status for the eventual outcome; do not automatically repeat the operation."));
    }

    private async Task<T> WaitWithDialogMonitoringAsync<T>(Task<T> completion, CancellationToken token)
    {
        while (!completion.IsCompleted)
        {
            await Task.WhenAny(completion, Task.Delay(TimeSpan.FromMilliseconds(500), token)).ConfigureAwait(false);
            if (completion.IsCompleted) break;
            token.ThrowIfCancellationRequested();
            _ = ReadDialogState();
        }
        return await completion.ConfigureAwait(false);
    }

    public ExcelVisibilityResult SetVisibility(bool visible)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
        lock (_applicationLease?.WindowControlLock ?? _windowControlLock)
        {
            var identity = _excelProcessIdentity
                ?? throw new InvalidOperationException("Excel process identity is unavailable.");
            var result = ExcelWindowVisibility.Set(identity, visible, _logger);
            UpdateConfirmedVisibility(result.IsVisible);
            _revealRequested = false;
            DialogVisibilityError = null;
            return result;
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Optional dialog reveal preserves the pending operation; failures are logged with cause and exposed in session blockingMessage, never reported as a successful reveal.")]
    private WorkbookRefreshState ReadDialogState()
    {
        var inspection = ExcelDialogProbe.Inspect(_excelProcessIdentity, _logger);
        lock (_applicationLease?.WindowControlLock ?? _windowControlLock)
        {
            if (inspection.State == WorkbookRefreshState.Ready)
            {
                _revealedDialog = 0;
                DialogVisibilityError = null;
            }
            else if (inspection.State == WorkbookRefreshState.DialogOpen &&
                _revealedDialog != inspection.Dialog && _disposed == 0 && !_operationTimedOut)
            {
                _revealedDialog = inspection.Dialog;
                try
                {
                    _ = SetVisibility(true);
                    _logger.LogWarning("{DialogGuidance}", ExcelBusyException.DialogGuidance);
                }
                catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException or ArgumentException)
                {
                    DialogVisibilityError = $"Excel's dialog was detected, but showing its window failed: {ex.Message}";
                    _logger.LogWarning(ex, "{DialogVisibilityError}", DialogVisibilityError);
                }
            }
        }
        return inspection.State;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Restoring an explicit hide after a deferred recovery reveal is cosmetic; failures are logged with cause and the following Visible read reports the window as shown, so the operation result is never misrepresented.")]
    private void UpdateVisibilitySnapshot(Exception? operationFailure)
    {
        var revision = _applicationLease?.VisibilityRevision;
        try
        {
            if (_revealRequested && _revealRequestRevision != revision)
                _revealRequested = false;
            if (_revealRequested && _excel != null)
            {
                _excel.Visible = true;
                // An explicit show/hide clears _revealRequested under the window lock. If one ran
                // while this deferred reveal was applied, restore that choice instead of undoing it.
                // The COM call stays outside the lock so native show/hide never waits on Excel.
                lock (_applicationLease?.WindowControlLock ?? _windowControlLock)
                {
                    if (_revealRequested && _applicationLease?.VisibilityRevision == revision)
                        _revealRequested = false;
                    else
                    {
                        _revealRequested = false;
                        if (!IsExcelVisible && _excelProcessIdentity is { } identity)
                        {
                            try
                            {
                                _ = ExcelWindowVisibility.Set(identity, visible: false, _logger);
                            }
                            catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException or ArgumentException)
                            {
                                // The operation's own result must stand; the visibility read below reports
                                // the window as still shown so the caller can hide it again.
                                _logger.LogWarning(ex, "Restoring the requested hidden Excel window failed");
                            }
                        }
                    }
                    revision = _applicationLease?.VisibilityRevision;
                }
            }
            var visible = _excel != null && (bool)_excel.Visible;
            BeforeVisibilitySnapshotPublicationHookForTests?.Invoke();
            UpdateConfirmedVisibility(visible, revision);
        }
        catch (COMException ex) when (ex.HResult == ResiliencePipelines.RPC_E_DISCONNECTED)
        {
            UpdateConfirmedVisibility(false, revision);
        }
        catch (InvalidComObjectException)
        {
            UpdateConfirmedVisibility(false, revision);
        }
        catch (Exception snapshotFailure) when (operationFailure is not null)
        {
            var failure = new AggregateException(
                "The Excel operation and updating its visibility snapshot both failed.",
                operationFailure, snapshotFailure);
            if (operationFailure is OperationCanceledException canceled)
            {
                throw new OperationCanceledException(canceled.Message, failure, canceled.CancellationToken);
            }

            throw failure;
        }
    }

    public void Save(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
        if (OperationStatus.ExecutionPending)
            throw new ExcelBusyException("Cannot save: Excel work is still pending. The workbook remains open.");
        ExcelBusyException.ThrowIfNotReady(GetExecutionPreflightState(), "save", this);
        Execute((ctx, ct) =>
        {
            if (!ExcelShutdownService.IsWorkbookOpen(_excel!, _workbook!))
            {
                throw new InvalidOperationException(
                    $"'{Path.GetFileName(_workbookPath)}' was closed in Excel. Nothing was saved. " +
                    "Close the session without saving to release it.");
            }
            ExcelBusyException.ThrowIfNotReady(
                ReadRefreshState(), "save", this);
            if (IsManagedTemporaryBatch)
            {
                foreach (var (path, workbook) in _workbooks!)
                {
                    if (!workbook.Saved)
                        ExcelShutdownService.SaveWorkbookWithTimeout(
                            workbook, Path.GetFileName(path), _logger, ct);
                }
                return 0;
            }
            ExcelShutdownService.SaveWorkbookWithTimeout(
                _workbook!,
                Path.GetFileName(_workbookPath),
                _logger,
                ct);
            return 0;
        }, cancellationToken);
    }

    /// <summary>Checks native dialogs and owned execution state without calling Excel COM.</summary>
    public WorkbookRefreshState GetExecutionPreflightState()
    {
        if (_disposed != 0)
        {
            return WorkbookRefreshState.Unknown;
        }
        // Window inspection must not queue behind COM work blocked by a modal prompt.
        var dialogState = ReadDialogState();
        if (dialogState != WorkbookRefreshState.Ready)
        {
            return dialogState;
        }
        if (!_startupCompletion.IsCompleted)
            return WorkbookRefreshState.Busy;
        if (_startupCompletion.IsFaulted)
            return WorkbookRefreshState.Unknown;
        if (OperationStatus.ExecutionPending)
        {
            return WorkbookRefreshState.Busy;
        }
        return WorkbookRefreshState.Ready;
    }

    public WorkbookRefreshState GetRefreshState() =>
        GetRefreshState(ReadRefreshState, selectionId: null);

    private WorkbookRefreshState GetRefreshState(
        Func<WorkbookRefreshState> readRefreshState,
        string? selectionId = null)
    {
        if (_operationTimedOut)
        {
            readRefreshState = ReadRefreshState;
            selectionId = null;
        }
        var preflight = GetExecutionPreflightState();
        if (preflight != WorkbookRefreshState.Ready)
            return preflight;

        if (!_readinessLock.TryEnter())
            return WorkbookRefreshState.Unknown;
        try
        {
            if (_readinessTask is { IsCompleted: false })
                return WorkbookRefreshState.Unknown;
            if (_readinessTask is { IsCompleted: true } previous)
            {
                _readinessTask = null;
                var state = previous.GetAwaiter().GetResult();
                if (string.Equals(_readinessSelectionId, selectionId, StringComparison.Ordinal)
                    && _readinessOperationStartedAtUtc == OperationStatus.StartedAtUtc)
                    return state;
            }
            var completion = new TaskCompletionSource<WorkbookRefreshState>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var inspectedOperation = OperationStatus.StartedAtUtc;
            var work = new ExcelWorkItem<WorkbookRefreshState>(
                () =>
                {
                    lock (_statusLock)
                        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
                    var state = _operationTimedOut ? ReadRefreshState() : readRefreshState();
                    if (state == WorkbookRefreshState.Ready)
                        _operationTimedOut = false;
                    return state;
                }, completion);
            if (!_workQueue.Writer.TryWrite(work))
                return WorkbookRefreshState.Unknown;
            _readinessTask = completion.Task;
            _readinessSelectionId = selectionId;
            _readinessOperationStartedAtUtc = inspectedOperation;
            try
            {
                // A started COM scan stays owned, but never holds a status request open.
                var state = completion.Task.WaitAsync(TimeSpan.FromMilliseconds(200)).GetAwaiter().GetResult();
                _readinessTask = null;
                return state;
            }
            catch (TimeoutException)
            {
                if (work.TryDiscard())
                {
                    completion.TrySetResult(WorkbookRefreshState.Unknown);
                    _logger.LogWarning("Excel refresh-state inspection expired in the queue; save and close remain blocked");
                }
                return WorkbookRefreshState.Unknown;
            }
        }
        finally
        {
            _readinessLock.Exit();
        }
    }

    private WorkbookRefreshState ReadRefreshState()
    {
        BeforeRefreshStateReadHookForTests?.Invoke();
        var dialogState = ReadDialogState();
        if (dialogState != WorkbookRefreshState.Ready) return dialogState;
        if (_attachedApplication is not null)
        {
            var discovery = RefreshAttachedWorkbooks();
            if (discovery.Errors.Count > 0)
                return WorkbookRefreshState.Unknown;
        }
        foreach (var workbook in _workbooks!.Values)
        {
            var state = ReadSelectedWorkbookRefreshState(workbook);
            if (state != WorkbookRefreshState.Ready) return state;
        }
        if (_revealRequested)
            UpdateVisibilitySnapshot(null);
        return WorkbookRefreshState.Ready;
    }

    private WorkbookRefreshState ReadAttachedWorkbookRefreshState(string workbookId)
    {
        BeforeRefreshStateReadHookForTests?.Invoke();
        var dialogState = ReadDialogState();
        if (dialogState != WorkbookRefreshState.Ready)
            return dialogState;

        var discovery = RefreshAttachedWorkbooks();
        var selected = discovery.Workbooks.FirstOrDefault(workbook =>
            string.Equals(workbook.WorkbookId, workbookId, StringComparison.Ordinal));
        if (selected is null)
            return discovery.Errors.Count == 0
                ? WorkbookRefreshState.Ready
                : WorkbookRefreshState.Unknown;
        return ReadSelectedWorkbookRefreshState(_workbooks![workbookId]);
    }

    private WorkbookRefreshState ReadSelectedWorkbookRefreshState(Excel.Workbook workbook)
    {
        // Protected View blocks active content and can reject connection inspection.
        if (FindProtectedViewAndRelease(workbook))
            return !_excel!.Ready || _excel.CalculationState == Excel.XlCalculationState.xlCalculating
                ? WorkbookRefreshState.Busy
                : WorkbookRefreshState.Ready;
        return WorkbookRefreshProbe.Read(_excel!, workbook);
    }

    public void Close()
    {
        lock (_disposeLock)
        {
            if (_disposed != 0)
            {
                if (!_teardownConfirmed)
                    CompleteDispose();
                return;
            }

            Execute((_, _) =>
            {
                ForgetWorkbooksClosedInExcel();
                ExcelBusyException.ThrowIfNotReady(ReadRefreshState(), "close", this);
                CloseOwnedWorkbooksBeforeShutdown();
                Interlocked.Exchange(ref _disposed, 1);
                _shutdownCts.Cancel();
                _workQueue.Writer.TryComplete();
            });
            CompleteDispose();
        }
    }

    internal void Shutdown()
    {
        lock (_disposeLock)
        {
            if (_disposed != 0)
            {
                if (!_teardownConfirmed)
                    CompleteDispose();
                return;
            }
            ExcelBusyException.ThrowIfNotReady(GetExecutionPreflightState(), "shutdown", this);
            RecoverTimedOutOperation();
            var shutdownWait = _operationTimeout < ComInteropConstants.ExcelQuitTimeout
                ? _operationTimeout
                : ComInteropConstants.ExcelQuitTimeout;
            ExcelBatchWait.Execute(this, shutdownWait, (_, _) =>
            {
                ForgetWorkbooksClosedInExcel();
                ExcelBusyException.ThrowIfNotReady(ReadRefreshState(), "shutdown", this);
                var alertsBeforeShutdown = _excel!.DisplayAlerts;
                try
                {
                    // Native prompts protect edits made after the saved-state check.
                    _excel.DisplayAlerts = true;
                    CloseOwnedWorkbooksBeforeShutdown(onlySaved: true);
                    if (_workbooks!.Count > 0)
                    {
                        _excel.Visible = true;
                        if (!_excel.Visible)
                            throw new InvalidOperationException("Cannot hand over unsaved workbooks: Excel visibility was not confirmed.");
                        UpdateConfirmedVisibility(true, _applicationLease!.VisibilityRevision);
                        if (_originalDisplayAlerts is { } displayAlerts)
                            _excel.DisplayAlerts = displayAlerts;
                        if (_excelProcessIdentity is not { } identity)
                            throw new InvalidOperationException("Cannot hand over unsaved workbooks without a confirmed Excel process identity.");
                        _leaveOwnedApplicationOpen = true;
                        _ownedProcessLeftOpen = true;
                        _logger.LogInformation("Leaving unsaved workbook(s) open and visible without saving or discarding changes");
                    }
                }
                catch (Exception shutdownFailure)
                {
                    try
                    {
                        _excel.DisplayAlerts = alertsBeforeShutdown;
                    }
                    catch (Exception restorationFailure)
                    {
                        throw new AggregateException(
                            "Excel shutdown failed and restoring DisplayAlerts also failed. The session remains connected.",
                            shutdownFailure, restorationFailure);
                    }
                    throw;
                }
                Interlocked.Exchange(ref _disposed, 1);
                _shutdownCts.Cancel();
                _workQueue.Writer.TryComplete();
                return 0;
            });
            CompleteDispose();
        }
    }

    internal string? FindOwnedWorkbookPath(OpenWorkbookInfo selected) =>
        ExcelBatchWait.Execute(this, TimeSpan.FromSeconds(2),
            (_, _) => FindOwnedWorkbookPathOnExcelThread(selected));

    internal string? FindOwnedWorkbookPathOnExcelThread(OpenWorkbookInfo selected)
    {
        if (_workbooks is null || _excel is null)
        {
            return null;
        }

        Excel.Workbooks? workbooks = null;
        try
        {
            workbooks = _excel.Workbooks;
            var count = workbooks.Count;
            for (var index = 1; index <= count; index++)
            {
                Excel.Workbook? current = null;
                try
                {
                    current = workbooks[index];
                    var tracked = _workbooks.FirstOrDefault(entry =>
                        ReferenceEquals(current, entry.Value));
                    if (tracked.Key is not null
                        && string.Equals(current.Name, selected.Name, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(GetWorkbookLocation(current.FullName), selected.Location,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return tracked.Key;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref current);
                }
            }
            return null;
        }
        finally
        {
            ComUtilities.Release(ref workbooks);
        }
    }

    internal bool DetachClosedOwnedWorkbook(string workbookPath)
    {
        var normalizedPath = WorkbookLocation.Normalize(workbookPath);
        lock (_disposeLock)
        {
            if (_disposed != 0 || _attachedApplication is not null)
                return false;

            var detached = Execute((_, _) =>
            {
                if (_workbooks is null || _excel is null || _workbooks.Count != 1)
                {
                    return false;
                }
                var tracked = _workbooks.FirstOrDefault(entry =>
                    string.Equals(WorkbookLocation.Normalize(entry.Key), normalizedPath,
                        StringComparison.OrdinalIgnoreCase));
                if (tracked.Key is null)
                    return false;

                Excel.Workbooks? workbooks = null;
                try
                {
                    workbooks = _excel.Workbooks;
                    var count = workbooks.Count;
                    for (var index = 1; index <= count; index++)
                    {
                        Excel.Workbook? current = null;
                        try
                        {
                            current = workbooks[index];
                            if (ReferenceEquals(current, tracked.Value))
                                return false;
                        }
                        finally
                        {
                            ComUtilities.Release(ref current);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref workbooks);
                }

                _leaveOwnedApplicationOpen = true;
                _ownedProcessLeftOpen = true;
                Interlocked.Exchange(ref _disposed, 1);
                _shutdownCts.Cancel();
                _workQueue.Writer.TryComplete();
                return true;
            });

            if (detached)
                CompleteDispose();
            return detached;
        }
    }

    private void CloseOwnedWorkbooksBeforeShutdown(bool onlySaved = false)
    {
        if (_excel is null || _workbooks is null)
            throw new InvalidOperationException("Owned Excel workbooks are not initialized.");

        var capabilities = _context?.Capabilities;
        foreach (var entry in _workbooks.ToArray())
        {
            var workbook = entry.Value;
            if (onlySaved && !workbook.Saved)
                continue;
            try
            {
                ExcelShutdownService.CloseWorkbookAndConfirm(
                    workbook, _excel, onlySaved ? null : false, entry.Key, _logger);
            }
            catch (Exception ex) when (_closedWorkbookPathsDuringShutdown.Count > 0)
            {
                var closedNames = _closedWorkbookPathsDuringShutdown
                    .Select(Path.GetFileName)
                    .ToArray();
                throw new InvalidOperationException(
                    $"Failed to close '{Path.GetFileName(entry.Key)}'. {ex.Message} " +
                    $"Earlier session workbooks already closed: {string.Join(", ", closedNames)}.",
                    ex);
            }

            _workbooks.Remove(entry.Key);
            _closedWorkbookPathsDuringShutdown.Add(entry.Key);
            ForgetTrackedWorkbook(entry.Value, capabilities);
        }
    }

    /// <summary>Releases tracked workbooks the user closed in Excel, without saving anything.</summary>
    private void ForgetWorkbooksClosedInExcel()
    {
        if (_excel is null || _workbooks is null)
            return;

        var capabilities = _context?.Capabilities;
        foreach (var entry in _workbooks.ToArray())
        {
            if (ExcelShutdownService.IsWorkbookOpen(_excel, entry.Value))
                continue;
            _logger.LogInformation("Workbook {FileName} was closed in Excel; releasing it without saving",
                Path.GetFileName(entry.Key));
            _workbooks.Remove(entry.Key);
            ForgetTrackedWorkbook(entry.Value, capabilities);
        }
    }

    /// <summary>
    /// Read-only check: true when every workbook this session opened was closed in Excel.
    /// Returns false at once when another thread is closing or shutting down this session.
    /// </summary>
    internal bool AreAllWorkbooksClosedInExcel(TimeSpan timeout)
    {
        if (!_disposeLock.TryEnter())
            return false;
        try
        {
            if (_disposed != 0 || _attachedApplication is not null
                || GetExecutionPreflightState() != WorkbookRefreshState.Ready)
                return false;
            return ExcelBatchWait.Execute(this, timeout, (_, _) =>
                _excel is not null && _workbooks is { Count: > 0 }
                && _workbooks.Values.All(workbook => !ExcelShutdownService.IsWorkbookOpen(_excel, workbook)));
        }
        finally
        {
            _disposeLock.Exit();
        }
    }

    private void ForgetTrackedWorkbook(Excel.Workbook workbook, ExcelCapabilities? capabilities)
    {
        if (ReferenceEquals(_workbook, workbook))
        {
            _workbook = null;
            _context = null;
        }

        Excel.Workbook? released = workbook;
        ComUtilities.Release(ref released);

        if (_workbook is null && _workbooks!.Count > 0)
        {
            var next = _workbooks.First();
            _workbookPath = next.Key;
            _allWorkbookPaths[0] = next.Key;
            _workbook = next.Value;
            _context = new ExcelContext(next.Key, _excel!, next.Value, capabilities)
            {
                ProcessIdentity = _excelProcessIdentity
            };
        }
    }

    internal IReadOnlyList<string> TakeClosedWorkbookPathsDuringShutdown()
    {
        var paths = _closedWorkbookPathsDuringShutdown.ToArray();
        _closedWorkbookPathsDuringShutdown.Clear();
        return paths;
    }

    internal bool IsManagedTemporaryBatch { get; set; }

    public void Dispose()
    {
        // The Service scope owns success cleanup and retains failures. Core's using
        // must neither discard partial edits nor replace the original exception.
        if (IsManagedTemporaryBatch && ExcelSession.HasTemporaryBatchOwner)
            return;

        lock (_disposeLock)
        {
            if (_disposed == 0 && IsExcelProcessAlive() &&
                (_operationTimedOut || !_startupCompletion.IsCompleted ||
                 OperationStatus.ExecutionPending || _readinessTask is { IsCompleted: false }))
            {
                RevealForRecovery();
                throw new ExcelBusyException("Cannot dispose: Excel work is still pending. The workbook remains open.");
            }
            if (_disposed == 0 && _startupCompletion.IsCompletedSuccessfully && IsExcelProcessAlive())
            {
                Close();
                return;
            }
            if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
            {
                if (!_teardownConfirmed)
                    CompleteDispose();
                return;
            }
            CompleteDispose();
        }
    }

    private void CompleteDispose()
    {
        lock (_statusLock)
        {
            _operationState = "closing";
            _operationStartedAt = System.Diagnostics.Stopwatch.GetTimestamp();
            _operationStartedAtUtc = DateTime.UtcNow;
            _completedElapsedSeconds = null;
        }
        _shutdownCts.Cancel();
        _workQueue.Writer.TryComplete();
        if (_staThread.IsAlive && !_staThread.Join(ComInteropConstants.StaThreadJoinTimeout))
        {
            RevealForRecovery(waitExpired: true);
            throw new TimeoutException(
                "The wait for Excel close/quit expired. The original close may still finish. " +
                $"The owned process has not been terminated. Excel window: {_windowReveal}.");
        }
        _teardownCompletion.Task.GetAwaiter().GetResult();

        if (_excelProcessIdentity is { } identity
            && (_attachedApplication is null || _quitEmptyOwnedApplication))
        {
            if (_ownedProcessLeftOpen)
            {
                if (!SessionManager.IsExcelProcessPreserved(identity))
                    SessionManager.PreserveExcelProcess(identity);
            }
            else if (!WaitForOwnedProcessExit(identity))
            {
                RevealForRecovery(waitExpired: true);
                throw new TimeoutException(
                    "Excel has not confirmed process exit after close/quit. " +
                    $"Its exact identity remains owned; it has not been terminated. Excel window: {_windowReveal}.");
            }
            else
            {
                SessionManager.UntrackExcelProcess(identity);
            }
        }
        FinishOperation(null);
        _teardownConfirmed = true;
        _shutdownCts.Dispose();
    }

    private static bool WaitForOwnedProcessExit(ExcelProcessIdentity identity)
    {
        var wait = System.Diagnostics.Stopwatch.StartNew();
        while (!OwnedProcessGuard.TryConfirmExited(identity) &&
               wait.Elapsed < ProcessTerminationPolicy.NormalGraceTimeout)
        {
            Thread.Sleep(50);
        }
        return OwnedProcessGuard.TryConfirmExited(identity);
    }

    internal static void FinalizeOwnedProcessTeardown(
        ExcelProcessIdentity identity,
        Func<ExcelProcessIdentity, bool> confirmExit)
    {
        ArgumentNullException.ThrowIfNull(confirmExit);
        if (!confirmExit(identity))
        {
            throw new InvalidOperationException(
                $"Excel process {identity.ProcessId} did not exit or its exit could not be confirmed. " +
                "Its exact process number and start time remain recorded. Inspect that Excel copy before deciding how to proceed.");
        }

        SessionManager.UntrackExcelProcess(identity);
    }

    internal static bool TryTerminateOwnedProcess(
        ExcelProcessIdentity identity,
        TimeSpan waitBeforeTermination,
        TimeSpan waitAfterTermination) =>
        OwnedProcessGuard.TryTerminate(
            identity,
            waitBeforeTermination,
            waitAfterTermination,
            out _);

    bool IExcelBatchTeardownState.TryConfirmOwnedProcessTeardown()
    {
        lock (_disposeLock)
        {
            if (_staThread.IsAlive
                || _excelProcessIdentity is not { } identity
                || !OwnedProcessGuard.TryConfirmExited(identity))
            {
                return false;
            }

            _applicationLease?.Dispose();
            _applicationLease = null;
            SessionManager.UntrackExcelProcess(identity);
            return true;
        }
    }

}
