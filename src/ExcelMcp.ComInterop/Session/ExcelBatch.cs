using System.Globalization;
using System.Reflection;
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
/// <item>For parallel processing, create multiple sessions for DIFFERENT files</item>
/// </list>
/// <para><b>Resource Cost:</b> Each ExcelBatch = one Excel.Application process (~50-100MB+ memory)</para>
/// </remarks>
internal sealed class ExcelBatch : IExcelBatch, IExcelBatchTeardownState, IExcelBatchRefreshState, IExcelBatchCloseState
{
    // P/Invoke for getting process ID from window handle
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private string _workbookPath; // Primary workbook path
    private readonly string[] _allWorkbookPaths; // All workbook paths (includes primary)
    private readonly bool _showExcel; // Whether to show Excel window
    private readonly bool _openReadOnly; // Whether existing workbooks are opened read-only
    private readonly bool _createNewFile; // Whether to create a new file instead of opening existing
    private readonly bool _isMacroEnabled; // For new files: whether to create .xlsm (macro-enabled)
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
    private string? _windowReveal;
    private Task _startupCompletion = Task.CompletedTask;
    private readonly TaskCompletionSource _teardownCompletion = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly object _statusLock = new();
    private readonly Lock _readinessLock = new();
    private bool _teardownConfirmed;
    private bool _staCleanupPending;
    private readonly AsyncLocal<(string? Command, CancellationToken CallerCancellation)> _commandContext = new();
    private string? _operationCommand;
    private long _operationStartedAt = System.Diagnostics.Stopwatch.GetTimestamp();
    private double? _completedElapsedSeconds;
    private string _operationState = "starting";
    private string? _operationError;
    private bool _waitExpired;
    private Task<WorkbookRefreshState>? _readinessTask;
    private bool _startupDetectedIrmProtectedWorkbook;

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
    internal Action? BeforeStaCleanupHookForTests { get; set; }

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
        bool deferStartup = false)
    {
        if (workbookPaths == null || workbookPaths.Length == 0)
            throw new ArgumentException("At least one workbook path is required", nameof(workbookPaths));

        // Reject known access failures before activating Excel; recheck at open for races.
        if (!createNewFile)
        {
            foreach (var path in workbookPaths)
            {
                if (!WorkbookLocation.IsRemote(path) && !FileAccessValidator.IsIrmProtected(path))
                    FileAccessValidator.ValidateFileNotLocked(path);
            }
        }

        _allWorkbookPaths = workbookPaths;
        _workbookPath = workbookPaths[0]; // Primary workbook
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
                                    if (_revealRequested)
                                        SessionManager.PreserveExcelProcess(_excelProcessIdentity.Value);
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

                // Workbook macro execution must remain available for explicit VBA operations on
                // reopened .xlsm sessions. Force-disabling macros at open makes vba.run impossible
                // later in the same batch. Validation opens are different: they inspect untrusted
                // workbooks and must always disable macros, even for .xlsm files.
                // msoAutomationSecurityLow = 1
                // msoAutomationSecurityForceDisable = 3
                // See: https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity
                // AutomationSecurity is typed as Office.MsoAutomationSecurity, an enum that lives in
                // office.dll (Microsoft.Office.Core). We do NOT reference or embed the Office.Core PIA,
                // so dispatch a plain integer without loading an Office.Core type.
                int automationSecurity = SelectAutomationSecurity(
                    _isMacroEnabled,
                    _openReadOnly,
                    _allWorkbookPaths);
                // Reflection dispatch avoids the dynamic binder's retained ITypeInfo RCW,
                // which can block STA termination after Excel.Quit has already returned.
                tempExcel.GetType().InvokeMember(
                    "AutomationSecurity", BindingFlags.SetProperty | BindingFlags.DoNotWrapExceptions, binder: null,
                    target: tempExcel, args: [automationSecurity],
                    culture: CultureInfo.InvariantCulture);

                // Open or create workbooks in the same Excel instance
                var tempWorkbooks = new Dictionary<string, Excel.Workbook>(StringComparer.OrdinalIgnoreCase);
                startupWorkbooks = tempWorkbooks;
                Excel.Workbook? primaryWorkbook = null;

                foreach (var path in _allWorkbookPaths)
                {
                    Excel.Workbook wb;
                    string normalizedPath = WorkbookLocation.Normalize(path);
                    bool isRemote = WorkbookLocation.IsRemote(normalizedPath);

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
                        // OPEN EXISTING FILE: Validate and open
                        bool isIrm = !isRemote && FileAccessValidator.IsIrmProtected(normalizedPath);

                        if (isIrm)
                        {
                            _startupDetectedIrmProtectedWorkbook = true;
                            if (!_showExcel)
                            {
                                throw new InvalidOperationException(CreateIrmRequiresVisibleSessionMessage(normalizedPath));
                            }

                            _logger.LogDebug(
                                "IRM-protected file detected: {FileName}. Excel determines editing permissions.",
                                Path.GetFileName(normalizedPath));
                        }
                        else if (!isRemote)
                        {
                            // CRITICAL: Check if file is locked at OS level BEFORE attempting Excel COM open
                            FileAccessValidator.ValidateFileNotLocked(path);
                        }

                        // Open workbook with Excel COM
                        Excel.Workbooks? workbooks = null;
                        try
                        {
                            BeforeWorkbookOpenHook?.Invoke(normalizedPath, _shutdownCts.Token);
                            _shutdownCts.Token.ThrowIfCancellationRequested();
                            workbooks = tempExcel.Workbooks;
                            // PIA gap: Open adds an LCID that can rewrite locale-specific table formats.
                            dynamic dispatch = (dynamic)(object)workbooks;
                            wb = (Excel.Workbook)dispatch.Open(
                                normalizedPath, UpdateLinks: 0, ReadOnly: _openReadOnly,
                                IgnoreReadOnlyRecommended: true, Notify: false, AddToMru: false);
                        }
                        finally
                        {
                            ComUtilities.Release(ref workbooks);
                        }
                    }

                    tempWorkbooks[normalizedPath] = wb;
                    AfterWorkbookOpenHookForTests?.Invoke(tempExcel, wb);
                    WorkbookLocation.ValidateOpenedWorkbookFormat(normalizedPath, wb.FileFormat);
                    if (isRemote && !wb.ReadOnly)
                    {
                        // Cloud AutoSave would otherwise persist edits before an explicit save
                        // and defeat close(save:false).
                        ExcelCapabilities.DisableAutoSave(() => wb.AutoSaveOn = false);
                    }
                    if (path == _workbookPath)
                    {
                        primaryWorkbook = wb;
                        startupPrimaryWorkbook = wb;
                    }
                }

                // All workbooks opened successfully — safe to apply user's visibility preference.
                // Enterprise auth/sign-in dialogs (if any) have already been dismissed.
                tempExcel.Visible = _showExcel || _revealRequested;
                _isExcelVisible = (bool)tempExcel.Visible;

                _excel = tempExcel;
                _workbook = primaryWorkbook;
                _workbooks = tempWorkbooks;
                _context = new ExcelContext(_workbookPath, _excel, _workbook!)
                {
                    ProcessIdentity = _excelProcessIdentity
                };

                FinishOperation(null);
                started.SetResult();

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
                    var cleanupExcel = _excel ?? startupExcel;
                    var cleanupWorkbook = _workbook ?? startupPrimaryWorkbook;
                    var cleanupWorkbooks = _workbooks ?? startupWorkbooks;

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

                    // Unified shutdown: use ExcelShutdownService for ALL workbook close/quit operations.
                    // Previously multi-workbook batches used bare COM calls without resilience,
                    // while single-workbook batches used ExcelShutdownService. Now both paths
                    // get the same exponential backoff retry for COM busy conditions.
                    if (cleanupWorkbooks != null && cleanupWorkbooks.Count > 1)
                    {
                        _logger.LogDebug("Closing {Count} workbooks via ExcelShutdownService", cleanupWorkbooks.Count);

                        // Close all non-primary workbooks first (without quitting Excel)
                        foreach (var kvp in cleanupWorkbooks.ToList())
                        {
                            if (kvp.Value == cleanupWorkbook)
                            {
                                continue; // Primary workbook closed last (with Quit)
                            }

                            // CloseAndQuit with excel=null closes workbook only, doesn't quit
                            ExcelShutdownService.CloseAndQuit(kvp.Value, null, false, kvp.Key, _logger);
                        }
                        cleanupWorkbooks.Clear();

                        // Close primary workbook AND quit Excel (with resilient retry)
                        ExcelShutdownService.CloseAndQuit(cleanupWorkbook, cleanupExcel, false, _workbookPath, _logger);
                    }
                    else
                    {
                        // Single workbook: same ExcelShutdownService path
                        ExcelShutdownService.CloseAndQuit(cleanupWorkbook, cleanupExcel, false, _workbookPath, _logger);
                    }

                    _workbook = null;
                    _excel = null;
                    _workbooks = null;
                    _context = null;
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

            if (_excelProcessIdentity is { } failedStartupIdentity)
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
        var protectedWorkbookHint = _startupDetectedIrmProtectedWorkbook
            ? $" {CreateIrmRequiresVisibleSessionMessage(_workbookPath)}"
            : string.Empty;

        return
            $"Excel startup timed out after {_startupTimeout.TotalSeconds} seconds while opening '{Path.GetFileName(_workbookPath)}'. " +
            "The workbook may be blocked on an interactive dialog, enterprise authentication, IRM/AIP prompt, external-link prompt, or an unresponsive open. " +
            $"The original startup remains owned and may still complete. Do not repeat the open/create to determine its outcome. " +
            $"Excel window: {_windowReveal}. Inspect session status for the eventual result.{protectedWorkbookHint}";
    }

    internal static int SelectAutomationSecurity(
        bool createsMacroEnabledWorkbook,
        bool openReadOnly,
        IReadOnlyCollection<string> workbookPaths)
    {
        if (openReadOnly)
        {
            return 3;
        }

        bool opensMacroEnabledWorkbook = createsMacroEnabledWorkbook ||
            workbookPaths.Any(path =>
                string.Equals(
                    Path.GetExtension(path),
                    ".xlsm",
                    StringComparison.OrdinalIgnoreCase));
        return opensMacroEnabledWorkbook ? 1 : 3;
    }

    private static string CreateIrmRequiresVisibleSessionMessage(string workbookPath)
    {
        return
            $"IRM/AIP-protected workbook '{Path.GetFileName(workbookPath)}' requires an interactive Excel session. " +
            "Retry with show=true so Excel stays visible for rights-management or enterprise-auth prompts, or open the workbook interactively first.";
    }

    public string WorkbookPath => _workbookPath;

    public void UpdateWorkbookPath(string workbookPath)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
        var normalizedPath = WorkbookLocation.Normalize(workbookPath);

        void UpdatePath()
        {
            if (_workbooks == null || _excel == null || _workbook == null)
            {
                throw new InvalidOperationException("Workbooks not initialized");
            }

            var previousPath = WorkbookLocation.Normalize(_workbookPath);
            if (!_workbooks.Remove(previousPath, out var workbook))
            {
                throw new InvalidOperationException($"Tracked workbook '{previousPath}' was not found.");
            }

            _workbooks[normalizedPath] = workbook;
            _workbookPath = normalizedPath;
            _allWorkbookPaths[0] = normalizedPath;
            _context = new ExcelContext(normalizedPath, _excel, _workbook, _context!.Capabilities)
            {
                ProcessIdentity = _excelProcessIdentity
            };
        }

        if (Thread.CurrentThread == _staThread)
            UpdatePath();
        else
            Execute((_, _) => UpdatePath());
    }

    public ILogger Logger => _logger;

    public int? ExcelProcessId => _excelProcessId;

    public TimeSpan OperationTimeout => _operationTimeout;

    public bool IsExcelVisible => _isExcelVisible;

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
                    Command = _operationCommand
                };
        }
    }

    internal void RevealForRecovery(bool waitExpired = false)
    {
        _revealRequested = true;
        if (_excelProcessIdentity is { } identity)
            SessionManager.PreserveExcelProcess(identity);
        lock (_statusLock)
        {
            if (waitExpired)
            {
                _waitExpired = true;
                if (_operationState is "starting" or "running" or "closing")
                    _operationTimedOut = true;
            }
        }
        var outcome = ExcelDialogProbe.Reveal(_excelProcessIdentity, _logger);
        lock (_statusLock)
            _windowReveal = outcome;
        if (outcome is "shown" or "alreadyVisible")
            _isExcelVisible = true;
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
                started = true;
                _operationStartedAt = System.Diagnostics.Stopwatch.GetTimestamp();
                _completedElapsedSeconds = null;
                _operationState = "running";
                _operationCommand = command;
                _operationError = null;
                _waitExpired = false;
                _windowReveal = null;
            }

            using var writeGuard = new ExcelWriteGuard((Excel.Application)_context!.App, _logger);

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
            return completion.Task.WaitAsync(timeoutCts?.Token ?? cancellationToken).GetAwaiter().GetResult();
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
        var dialogState = ExcelDialogProbe.ReadWithText(_excelProcessIdentity, _logger, out var dialogText);
        var dialogOpen = dialogState is WorkbookRefreshState.DialogOpen or WorkbookRefreshState.NativeQueryApproval;
        return new TimeoutException(
            $"Excel operation timed out after {_operationTimeout.TotalSeconds} seconds for '{Path.GetFileName(_workbookPath)}'. " +
            $"The workbook remains open and the operation may still change it. No changes were saved or discarded by timeout recovery. Excel window: {_windowReveal}. " +
            (dialogOpen
                ? (dialogState == WorkbookRefreshState.NativeQueryApproval
                    ? "Excel was waiting for approval to run a native database query. "
                    : "Excel had a dialog open when the operation timed out and may have been waiting for input. ") +
                  "This timeout does not itself cancel the operation or undo workbook changes. " +
                  "The captured dialog text is included below. " +
                  $"Dialog text: {dialogText}"
                : "Excel may be unresponsive or the operation is taking longer than expected. " +
                  "Inspect session status for the eventual outcome; do not automatically repeat the operation."));
    }

    private void UpdateVisibilitySnapshot(Exception? operationFailure)
    {
        try
        {
            if (_revealRequested && _excel != null)
            {
                _excel.Visible = true;
                _revealRequested = false;
            }
            _isExcelVisible = _excel != null && (bool)_excel.Visible;
        }
        catch (COMException ex) when (ex.HResult == ResiliencePipelines.RPC_E_DISCONNECTED)
        {
            _isExcelVisible = false;
        }
        catch (InvalidComObjectException)
        {
            _isExcelVisible = false;
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
        if (OperationStatus.ExecutionPending)
            throw new ExcelBusyException("Cannot save: Excel work is still pending. The workbook remains open.");
        Execute((ctx, ct) =>
        {
            ExcelBusyException.ThrowIfNotReady(
                ReadRefreshState(), "save");
            ExcelShutdownService.SaveWorkbookWithTimeout(
                _workbook!,
                Path.GetFileName(_workbookPath),
                _logger,
                ct);
            return 0;
        }, cancellationToken);
    }

    internal WorkbookRefreshState GetExecutionPreflightState()
    {
        if (_disposed != 0)
        {
            return WorkbookRefreshState.Unknown;
        }
        // Window inspection must not queue behind COM work blocked by a modal prompt.
        var dialogState = ExcelDialogProbe.Read(_excelProcessIdentity, _logger);
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

    public WorkbookRefreshState GetRefreshState()
    {
        var preflight = GetExecutionPreflightState();
        if (preflight != WorkbookRefreshState.Ready)
            return preflight;

        lock (_readinessLock)
        {
            if (_readinessTask is { IsCompleted: false })
                return WorkbookRefreshState.Unknown;
            if (_readinessTask is { IsCompleted: true } previous)
            {
                _readinessTask = null;
                var state = previous.GetAwaiter().GetResult();
                if (state != WorkbookRefreshState.Ready || !_operationTimedOut)
                    return state;
            }
            var completion = new TaskCompletionSource<WorkbookRefreshState>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            var work = new ExcelWorkItem<WorkbookRefreshState>(
                () =>
                {
                    var state = ReadRefreshState();
                    if (state == WorkbookRefreshState.Ready)
                        _operationTimedOut = false;
                    return state;
                }, completion);
            if (!_workQueue.Writer.TryWrite(work))
                return WorkbookRefreshState.Unknown;
            _readinessTask = completion.Task;
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
    }

    private WorkbookRefreshState ReadRefreshState()
    {
        BeforeRefreshStateReadHookForTests?.Invoke();
        var dialogState = ExcelDialogProbe.Read(_excelProcessIdentity, _logger);
        if (dialogState != WorkbookRefreshState.Ready) return dialogState;
        foreach (var workbook in _workbooks!.Values)
        {
            var state = WorkbookRefreshProbe.Read(_excel!, workbook);
            if (state != WorkbookRefreshState.Ready) return state;
        }
        if (_revealRequested)
            UpdateVisibilitySnapshot(null);
        return WorkbookRefreshState.Ready;
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
                ExcelBusyException.ThrowIfNotReady(ReadRefreshState(), "close");
                // Commit shutdown on the STA only after readiness is confirmed.
                Interlocked.Exchange(ref _disposed, 1);
                _shutdownCts.Cancel();
                _workQueue.Writer.TryComplete();
            });
            CompleteDispose();
        }
    }

    public void Dispose()
    {
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

        if (_excelProcessIdentity is { } identity)
        {
            if (!WaitForOwnedProcessExit(identity))
            {
                RevealForRecovery(waitExpired: true);
                throw new TimeoutException(
                    "Excel has not confirmed process exit after close/quit. " +
                    $"Its exact identity remains owned; it has not been terminated. Excel window: {_windowReveal}.");
            }
            SessionManager.UntrackExcelProcess(identity);
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
                "Its exact PID/start-time identity remains tracked for later pipe-scoped cleanup.");
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
        if (_excelProcessIdentity is not { } identity
            || !OwnedProcessGuard.TryConfirmExited(identity))
        {
            return false;
        }

        SessionManager.UntrackExcelProcess(identity);
        return true;
    }

}
