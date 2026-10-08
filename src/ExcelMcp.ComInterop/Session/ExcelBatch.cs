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
    private int _executingWorkItem;
    private int? _excelProcessId; // Excel.exe process ID for force-kill if needed
    private ExcelProcessIdentity? _excelProcessIdentity;
    private volatile bool _isExcelVisible;
    private bool _operationTimedOut; // Track if an operation timed out for aggressive cleanup
    private bool _startupDetectedIrmProtectedWorkbook;

    /// <summary>
    /// When true, suppresses the "start visible during open" behavior.
    /// Used by test infrastructure to avoid flashing Excel windows during automated test runs.
    /// Production code should never set this.
    /// </summary>
    internal static bool SuppressVisibleDuringOpen { get; set; }

    /// <summary>
    /// Test-only seam for simulating a blocked workbook open path during startup.
    /// Production code must leave this null.
    /// </summary>
    internal static Action<string, CancellationToken>? BeforeWorkbookOpenHook { get; set; }
    internal static Action<object, object>? AfterWorkbookOpenHookForTests { get; set; }
    internal static Func<int, ExcelProcessIdentity?>? TrackProcessIdentityHookForTests { get; set; }

    internal static Func<ExcelProcessIdentity, bool>? FailedStartupTerminationHook { get; set; }

    internal static Func<ExcelProcessIdentity, bool>? FailedStartupExitConfirmationHook { get; set; }

    internal static Action? WorkItemQueuedHookForTests { get; set; }
    internal Action? BeforeRefreshStateReadHookForTests { get; set; }

    // COM state (STA thread only)
    private Excel.Application? _excel;
    private Excel.Workbook? _workbook; // Primary workbook
    private Dictionary<string, Excel.Workbook>? _workbooks; // All workbooks keyed by normalized path
    private ExcelContext? _context;

    private interface IExcelWorkItem
    {
        bool IsExecuting { get; }

        bool TryExecute();

        bool TryDiscard(Exception? exception = null);
    }

    private sealed class ExcelWorkItem<T>(
        Func<T> operation,
        TaskCompletionSource<T> completion) : IExcelWorkItem
    {
        private const int Queued = 0;
        private const int Executing = 1;
        private const int Completed = 2;
        private const int Discarded = 3;
        private int _state;

        public bool IsExecuting => Volatile.Read(ref _state) == Executing;

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
                completion.TrySetResult(result);
            }
            catch (OperationCanceledException ex)
            {
                Volatile.Write(ref _state, Completed);
                completion.TrySetCanceled(ex.CancellationToken);
            }
            catch (Exception ex)
            {
                Volatile.Write(ref _state, Completed);
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
        TimeSpan? startupTimeout = null)
        : this(
            workbookPaths,
            logger,
            show,
            openReadOnly,
            createNewFile: false,
            isMacroEnabled: false,
            operationTimeout: operationTimeout,
            startupTimeout: startupTimeout)
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
    internal static ExcelBatch CreateNewWorkbook(string filePath, bool isMacroEnabled, ILogger<ExcelBatch>? logger = null, bool show = false, TimeSpan? operationTimeout = null)
    {
        return new ExcelBatch(
            [filePath],
            logger,
            show,
            openReadOnly: false,
            createNewFile: true,
            isMacroEnabled: isMacroEnabled,
            operationTimeout: operationTimeout);
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
        TimeSpan? startupTimeout = null)
    {
        if (workbookPaths == null || workbookPaths.Length == 0)
            throw new ArgumentException("At least one workbook path is required", nameof(workbookPaths));

        _allWorkbookPaths = workbookPaths;
        _workbookPath = workbookPaths[0]; // Primary workbook
        _showExcel = show;
        _openReadOnly = openReadOnly;
        _createNewFile = createNewFile;
        _isMacroEnabled = isMacroEnabled;
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

        _staThread = new Thread(() =>
        {
            Excel.Application? startupExcel = null;
            Excel.Workbook? startupPrimaryWorkbook = null;
            Dictionary<string, Excel.Workbook>? startupWorkbooks = null;
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
                    tempExcel = (Excel.Application)Activator.CreateInstance(excelType)!;
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
                // SuppressVisibleDuringOpen: test infrastructure disables this to avoid
                // flashing windows during automated runs.
                if (!SuppressVisibleDuringOpen)
                {
                    tempExcel.Visible = true;
                }
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
                    _logger.LogWarning(ex, "Failed to capture Excel process identity during startup.");
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
                            // The Excel PIA adds an LCID to Open, which makes Excel rewrite
                            // locale-specific table column format definitions during a later save.
                            // IDispatch preserves the workbook's native definitions.
                            workbooks = tempExcel.Workbooks;
                            dynamic workbooksDispatch = (dynamic)(object)workbooks;
                            // Protection does not imply read-only rights. Excel enforces the
                            // signed-in user's permissions; only validation forces read-only.
                            wb = (Excel.Workbook)workbooksDispatch.Open(
                                normalizedPath,
                                UpdateLinks: 0,
                                ReadOnly: _openReadOnly,
                                IgnoreReadOnlyRecommended: true,
                                Notify: false,
                                AddToMru: false);
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
                tempExcel.Visible = _showExcel;
                _isExcelVisible = (bool)tempExcel.Visible;

                _excel = tempExcel;
                _workbook = primaryWorkbook;
                _workbooks = tempWorkbooks;
                _context = new ExcelContext(_workbookPath, _excel, _workbook!);

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

                            Interlocked.Exchange(ref _executingWorkItem, 1);
                            try
                            {
                                work.TryExecute();
                            }
                            finally
                            {
                                Interlocked.Exchange(ref _executingWorkItem, 0);
                            }
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
            }
        })
        {
            IsBackground = true,
            Name = $"ExcelBatch-{Path.GetFileName(_workbookPath)}"
        };

        // CRITICAL: Set STA apartment state before starting thread
        _staThread.SetApartmentState(ApartmentState.STA);
        _staThread.Start();

        // Wait for STA thread to initialize. If startup fails, the STA thread may still be in
        // its finally-block cleanup path (closing a hidden Excel instance created before the
        // failure). Wait for that cleanup before rethrowing so callers do not immediately race
        // a lingering hidden Excel.exe from the failed startup attempt.
        try
        {
            bool completedInTime;
            try
            {
                completedInTime = started.Task.Wait(_startupTimeout);
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
                _workQueue.Writer.TryComplete();
                _shutdownCts.Cancel();

                throw new TimeoutException(CreateStartupTimeoutMessage());
            }

            started.Task.GetAwaiter().GetResult();
        }
        catch (Exception startupFailure)
        {
            if (_staThread.IsAlive)
            {
                _ = _staThread.Join(TimeSpan.FromSeconds(10));
            }

            if (_excelProcessIdentity is { } failedStartupIdentity)
            {
                try
                {
                    FinalizeFailedStartupOwnedProcess(
                        failedStartupIdentity,
                        FailedStartupTerminationHook
                        ?? (identity => TryTerminateOwnedProcess(
                            identity,
                            TimeSpan.FromSeconds(5),
                            TimeSpan.FromSeconds(3))),
                        FailedStartupExitConfirmationHook
                        ?? OwnedProcessGuard.TryConfirmExited);
                }
                catch (Exception teardownFailure)
                {
                    throw new InvalidOperationException(
                        $"Excel startup failed and exact process teardown could not be confirmed for " +
                        $"process {failedStartupIdentity.ProcessId}. The process identity remains tracked " +
                        "for ProcessExit cleanup.",
                        new AggregateException(startupFailure, teardownFailure));
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
            "Corrective action: retry the file open/create with a larger timeout_seconds value (CLI: --timeout-seconds <seconds>) if the workbook is just slow, " +
            $"or retry with show=true (CLI: --show) so Excel is visible for prompts.{protectedWorkbookHint}";
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

        Execute((_, _) =>
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
            _context = new ExcelContext(normalizedPath, _excel, _workbook, _context!.Capabilities);
        });
    }

    public ILogger Logger => _logger;

    public int? ExcelProcessId => _excelProcessId;

    public TimeSpan OperationTimeout => _operationTimeout;

    public bool IsExcelVisible => _isExcelVisible;

    public bool HasTimedOutOperation => _operationTimedOut;

    public bool IsExcelProcessAlive()
    {
        if (_disposed != 0) return false;
        if (!_excelProcessId.HasValue)
        {
            // PID capture is best-effort during startup. Under load, Excel can still be healthy
            // even when Hwnd-based PID discovery misses the short retry window. Treat this as
            // "unknown but assumed alive" so callers do not tear down a live session before the
            // first COM command runs; real COM failures still surface on the actual operation.
            return true;
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

        // Fail fast if a previous operation timed out or was cancelled while the STA thread
        // was stuck in IDispatch.Invoke. The STA thread cannot process new work items until
        // the hung COM call returns (which may be never). Without this check, new callers
        // would queue work and block until their own timeout expires.
        if (_operationTimedOut)
        {
            throw new TimeoutException(
                $"A previous operation timed out or was cancelled for '{Path.GetFileName(_workbookPath)}'. " +
                "The Excel COM thread may be unresponsive. Please close this session and create a new one.");
        }

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
        var workItem = new ExcelWorkItem<T>(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            using var writeGuard = new ExcelWriteGuard((Excel.Application)_context!.App, _logger);

            try
            {
                return operation(_context!, cancellationToken);
            }
            finally
            {
                UpdateVisibilitySnapshot();
            }
        }, completion);

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
        try
        {
            if (cancellationToken.CanBeCanceled)
            {
                // Caller controls the timeout — use their token exclusively
                return completion.Task.WaitAsync(cancellationToken).GetAwaiter().GetResult();
            }
            else
            {
                // No caller timeout — apply session-level operation timeout as safety net
                using var timeoutCts = new CancellationTokenSource(_operationTimeout);
                return completion.Task.WaitAsync(timeoutCts.Token).GetAwaiter().GetResult();
            }
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            // Session timeout occurred (not caller cancellation) — only happens in the else branch
            var expiredWhileQueued = workItem.TryDiscard();
            if (!expiredWhileQueued && workItem.IsExecuting)
            {
                _operationTimedOut = true;
            }

            if (expiredWhileQueued)
            {
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
            throw new TimeoutException(
                expiredWhileQueued
                    ? $"Excel operation expired in the session queue after {_operationTimeout.TotalSeconds} seconds for '{Path.GetFileName(_workbookPath)}' and was not executed."
                    : $"Excel operation timed out after {_operationTimeout.TotalSeconds} seconds for '{Path.GetFileName(_workbookPath)}'. " +
                      "Excel may be unresponsive or the operation is taking longer than expected. " +
                      "Consider increasing timeoutSeconds when opening the session.");
        }
        catch (OperationCanceledException)
        {
            var cancelledWhileQueued = workItem.TryDiscard();
            _logger.LogDebug("Operation cancelled or timed out for {FileName}", Path.GetFileName(_workbookPath));
            if (!cancelledWhileQueued && workItem.IsExecuting)
            {
                _operationTimedOut = true;
            }

            throw;
        }
    }

    private void UpdateVisibilitySnapshot()
    {
        try
        {
            _isExcelVisible = _excel != null && (bool)_excel.Visible;
        }
        catch (COMException)
        {
            _isExcelVisible = false;
        }
        catch (InvalidComObjectException)
        {
            _isExcelVisible = false;
        }
    }

    public void Save(CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ExcelBusyException.ThrowIfNotReady(GetRefreshState(), "save");
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

    public WorkbookRefreshState GetRefreshState()
    {
        if (_disposed != 0 || _operationTimedOut)
        {
            return WorkbookRefreshState.Unknown;
        }
        // Window inspection must not queue behind COM work blocked by a modal prompt.
        var dialogState = ExcelDialogProbe.Read(_excelProcessIdentity, _logger);
        if (dialogState != WorkbookRefreshState.Ready)
        {
            return dialogState;
        }
        if (Volatile.Read(ref _executingWorkItem) != 0)
        {
            return WorkbookRefreshState.Busy;
        }

        var completion = new TaskCompletionSource<WorkbookRefreshState>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        var work = new ExcelWorkItem<WorkbookRefreshState>(
            ReadRefreshState, completion);
        if (!_workQueue.Writer.TryWrite(work))
        {
            return WorkbookRefreshState.Unknown;
        }
        try
        {
            // Keep queued probes short without imposing the same deadline on a started COM scan.
            return completion.Task.WaitAsync(TimeSpan.FromSeconds(1)).GetAwaiter().GetResult();
        }
        catch (TimeoutException)
        {
            if (work.TryDiscard())
            {
                completion.TrySetResult(WorkbookRefreshState.Unknown);
                _logger.LogWarning("Excel refresh-state inspection expired in the queue; save and close remain blocked");
                return WorkbookRefreshState.Unknown;
            }
            try
            {
                return completion.Task.WaitAsync(_operationTimeout).GetAwaiter().GetResult();
            }
            catch (TimeoutException)
            {
                _logger.LogWarning(
                    "Started Excel refresh-state inspection did not complete within {Timeout}; save and close remain blocked",
                    _operationTimeout);
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
            var state = WorkbookRefreshProbe.Read(_excel!, workbook, _logger);
            if (state != WorkbookRefreshState.Ready) return state;
        }
        return WorkbookRefreshState.Ready;
    }

    public void Close()
    {
        lock (_disposeLock)
        {
            if (_disposed != 0) return;
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
            if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
            {
                _logger.LogDebug("[Thread {CallingThread}] Dispose skipped - already disposed for {FileName}",
                    Environment.CurrentManagedThreadId, Path.GetFileName(_workbookPath));
                return;
            }
            CompleteDispose();
        }
    }

    private void CompleteDispose()
    {
        var callingThread = Environment.CurrentManagedThreadId;

        _logger.LogDebug("[Thread {CallingThread}] Dispose starting for {FileName}", callingThread, Path.GetFileName(_workbookPath));

        // Cancel the shutdown token FIRST to wake up the message pump
        _logger.LogDebug("[Thread {CallingThread}] Cancelling shutdown token for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _shutdownCts.Cancel();

        // Then complete the work queue
        _logger.LogDebug("[Thread {CallingThread}] Completing work queue for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _workQueue.Writer.TryComplete();

        _logger.LogDebug("[Thread {CallingThread}] Waiting for STA thread (Id={STAThread}) to exit for {FileName}", callingThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));

        // When operation timed out, the STA thread is stuck in IDispatch.Invoke (unmanaged COM call
        // that cannot be cancelled). Kill the Excel process FIRST to unblock the STA thread, then wait.
        if (_operationTimedOut
            && _excelProcessIdentity is { } timedOutIdentity
            && _staThread != null
            && _staThread.IsAlive)
        {
            // INSTRUMENTATION: Track pre-emptive kill path entry
            _logger.LogWarning(
                "[DIAG-DISPOSE-TIMEOUT-PREKILL] [Thread {CallingThread}] Operation timed out — force-killing Excel process {ProcessId} BEFORE waiting for STA thread to unblock IDispatch.Invoke for {FileName}",
                callingThread, timedOutIdentity.ProcessId, Path.GetFileName(_workbookPath));
            SessionDiagnostics.WriteStdErr(
                $"[DIAG-DISPOSE-TIMEOUT-PREKILL] [Thread {callingThread}] Operation timed out — force-killing Excel process {timedOutIdentity.ProcessId} BEFORE waiting for STA thread to unblock IDispatch.Invoke for {Path.GetFileName(_workbookPath)}");
            if (OwnedProcessGuard.TryTerminate(
                    timedOutIdentity,
                    TimeSpan.Zero,
                    TimeSpan.FromSeconds(5),
                    out var preemptivelyTerminated))
            {
                if (preemptivelyTerminated)
                {
                    // INSTRUMENTATION: Track successful pre-emptive kill
                    _logger.LogInformation(
                        "[DIAG-DISPOSE-TIMEOUT-PREKILL-SUCCESS] [Thread {CallingThread}] Force-killed Excel process {ProcessId} (pre-emptive, before STA join)",
                        callingThread, timedOutIdentity.ProcessId);
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-TIMEOUT-PREKILL-SUCCESS] [Thread {callingThread}] Force-killed Excel process {timedOutIdentity.ProcessId} (pre-emptive, before STA join)");
                }
                else
                {
                    _logger.LogDebug(
                        "[DIAG-DISPOSE-TIMEOUT-PREKILL-ALREADY-GONE] [Thread {CallingThread}] Excel process {ProcessId} already exited or its PID was reused",
                        callingThread, timedOutIdentity.ProcessId);
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-TIMEOUT-PREKILL-ALREADY-GONE] [Thread {callingThread}] Excel process {timedOutIdentity.ProcessId} already exited or its PID was reused");
                }
            }
            else
            {
                _logger.LogWarning("[DIAG-DISPOSE-TIMEOUT-PREKILL-FAILED] [Thread {CallingThread}] Failed to force-kill Excel process {ProcessId}", callingThread, timedOutIdentity.ProcessId);
                SessionDiagnostics.WriteStdErr(
                    $"[DIAG-DISPOSE-TIMEOUT-PREKILL-FAILED] [Thread {callingThread}] Failed to force-kill Excel process {timedOutIdentity.ProcessId}");
            }
        }

        // Wait for STA thread to finish cleanup (with timeout)
        var staExitedWithoutForce = false;
        if (_staThread != null && _staThread.IsAlive)
        {
            // Use shorter timeout if operation timed out (Excel is likely hung / already killed above)
            var joinTimeout = _operationTimedOut
                ? TimeSpan.FromSeconds(10)  // Aggressive: 10 seconds when operation timed out
                : ComInteropConstants.StaThreadJoinTimeout;  // Normal: 45 seconds

            var reasonSuffix = _operationTimedOut ? " (operation timed out - aggressive cleanup)" : "";
            _logger.LogDebug(
                "[Thread {CallingThread}] Calling Join() with {Timeout} timeout on STA={STAThread}, file={FileName}{Reason}",
                callingThread, joinTimeout, _staThread.ManagedThreadId, Path.GetFileName(_workbookPath), reasonSuffix);

            // CRITICAL: StaThreadJoinTimeout >= ExcelQuitTimeout + margin (currently 45 seconds total).
            // The join must wait at least as long as CloseAndQuit() can take, otherwise Dispose() returns
            // before Excel has finished closing, causing "file still open" issues in subsequent operations.
            if (!_staThread.Join(joinTimeout))
            {
                // STA thread didn't exit - Excel cleanup is severely stuck
                var reasonForError = _operationTimedOut ? " (operation previously timed out)" : "";
                // INSTRUMENTATION: Track STA join timeout (the key failure mode)
                _logger.LogError(
                    "[DIAG-DISPOSE-STA-JOIN-TIMEOUT] [Thread {CallingThread}] STA thread (Id={STAThread}) did NOT exit within {Timeout} for {FileName}. " +
                    "Excel cleanup is severely stuck{Reason}. Attempting force-kill.",
                    callingThread, _staThread.ManagedThreadId, joinTimeout, Path.GetFileName(_workbookPath), reasonForError);
                SessionDiagnostics.WriteStdErr(
                    $"[DIAG-DISPOSE-STA-JOIN-TIMEOUT] [Thread {callingThread}] STA thread (Id={_staThread.ManagedThreadId}) did NOT exit within {joinTimeout} for {Path.GetFileName(_workbookPath)}. Excel cleanup is severely stuck{reasonForError}. Attempting force-kill.");

                // Force-kill only the exact Excel identity captured for this batch.
                if (_excelProcessIdentity is { } stuckIdentity)
                {
                    _logger.LogWarning(
                        "[DIAG-DISPOSE-FORCE-KILL-ATTEMPT] [Thread {CallingThread}] Force-killing Excel process {ProcessId} for {FileName}",
                        callingThread, stuckIdentity.ProcessId, Path.GetFileName(_workbookPath));
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-FORCE-KILL-ATTEMPT] [Thread {callingThread}] Force-killing Excel process {stuckIdentity.ProcessId} for {Path.GetFileName(_workbookPath)}");

                    if (OwnedProcessGuard.TryTerminate(
                            stuckIdentity,
                            TimeSpan.Zero,
                            TimeSpan.FromSeconds(5),
                            out var forceTerminated))
                    {
                        var outcome = forceTerminated
                            ? "force-killed"
                            : "already exited or PID was reused";
                        _logger.LogInformation(
                            "[DIAG-DISPOSE-FORCE-KILL-SUCCESS] [Thread {CallingThread}] Excel process {ProcessId} was {Outcome}",
                            callingThread, stuckIdentity.ProcessId, outcome);
                        SessionDiagnostics.WriteStdErr(
                            $"[DIAG-DISPOSE-FORCE-KILL-SUCCESS] [Thread {callingThread}] Excel process {stuckIdentity.ProcessId} was {outcome}");

                        if (_staThread.Join(TimeSpan.FromSeconds(5)))
                        {
                            _logger.LogDebug("[DIAG-DISPOSE-STA-EXIT-AFTER-KILL] [Thread {CallingThread}] STA thread exited after force-kill", callingThread);
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-DISPOSE-STA-EXIT-AFTER-KILL] [Thread {callingThread}] STA thread exited after force-kill");
                        }
                        else
                        {
                            _logger.LogWarning(
                                "[DIAG-DISPOSE-STA-LEAK] [Thread {CallingThread}] STA thread still stuck even after force-kill. Thread leak.",
                                callingThread);
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-DISPOSE-STA-LEAK] [Thread {callingThread}] STA thread still stuck even after force-kill. Thread leak.");
                        }
                    }
                    else
                    {
                        _logger.LogError(
                            "[Thread {CallingThread}] Failed to force-kill Excel process {ProcessId}",
                            callingThread, stuckIdentity.ProcessId);
                    }
                }
                else
                {
                    _logger.LogError(
                        "[Thread {CallingThread}] No Excel process identity captured - cannot force-kill. Process will leak.",
                        callingThread);
                }
            }
            else
            {
                staExitedWithoutForce = true;
            }
        }
        else
        {
            staExitedWithoutForce = true;
            _logger.LogDebug("[Thread {CallingThread}] STA thread was null or not alive for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        }

        try
        {
            // Wait for Excel process to fully terminate to prevent CO_E_SERVER_EXEC_FAILURE
            // on subsequent Activator.CreateInstance calls. excel.Quit() + COM release doesn't
            // guarantee the EXCEL.EXE process has exited — rapid create/destroy cycles can fail.
            if (_excelProcessIdentity is { } lingeringIdentity)
            {
                _logger.LogDebug(
                    "[DIAG-DISPOSE-PROCESS-WAIT] [Thread {CallingThread}] Waiting for Excel process {ProcessId} to exit for {FileName}",
                    callingThread, lingeringIdentity.ProcessId, Path.GetFileName(_workbookPath));
                SessionDiagnostics.WriteStdErr(
                    $"[DIAG-DISPOSE-PROCESS-WAIT] [Thread {callingThread}] Waiting for Excel process {lingeringIdentity.ProcessId} to exit for {Path.GetFileName(_workbookPath)}");

                var lingeringProcessTerminated = false;
                var normalShutdown = staExitedWithoutForce && !_operationTimedOut;
                FinalizeOwnedProcessTeardown(
                    lingeringIdentity,
                    identity => OwnedProcessGuard.TryTerminate(
                        identity,
                        normalShutdown ? ProcessTerminationPolicy.NormalGraceTimeout : TimeSpan.FromSeconds(5),
                        normalShutdown ? ProcessTerminationPolicy.NormalForcedExitTimeout : ProcessTerminationPolicy.ProcessExitTimeout,
                        out lingeringProcessTerminated,
                        overallTimeout: normalShutdown ? ProcessTerminationPolicy.NormalShutdownBudget : null));

                if (lingeringProcessTerminated)
                {
                    _logger.LogInformation(
                        "[DIAG-DISPOSE-PROCESS-LINGER-KILLED] [Thread {CallingThread}] Force-killed lingering Excel process {ProcessId} for {FileName}",
                        callingThread, lingeringIdentity.ProcessId, Path.GetFileName(_workbookPath));
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-PROCESS-LINGER-KILLED] [Thread {callingThread}] Force-killed lingering Excel process {lingeringIdentity.ProcessId} for {Path.GetFileName(_workbookPath)}");
                }
                else
                {
                    _logger.LogDebug(
                        "[DIAG-DISPOSE-PROCESS-EXITED] [Thread {CallingThread}] Excel process {ProcessId} exited normally or its PID was reused for {FileName}",
                        callingThread, lingeringIdentity.ProcessId, Path.GetFileName(_workbookPath));
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-PROCESS-EXITED] [Thread {callingThread}] Excel process {lingeringIdentity.ProcessId} exited normally or its PID was reused for {Path.GetFileName(_workbookPath)}");
                }
            }

            _logger.LogDebug("[Thread {CallingThread}] Dispose COMPLETED for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        }
        catch (InvalidOperationException ex) when (_excelProcessIdentity is { })
        {
            _logger.LogError(
                ex,
                "[DIAG-DISPOSE-PROCESS-LINGER-KILL-FAILED] [Thread {CallingThread}] Excel teardown failed for {FileName}; the exact process identity remains tracked",
                callingThread,
                Path.GetFileName(_workbookPath));
            SessionDiagnostics.WriteStdErr(
                $"[DIAG-DISPOSE-PROCESS-LINGER-KILL-FAILED] [Thread {callingThread}] {ex.Message}");
            throw;
        }
        finally
        {
            // COM cleanup runs on the STA thread; this local resource must also be
            // released even when final process termination cannot be confirmed.
            _logger.LogDebug("[Thread {CallingThread}] Disposing CancellationTokenSource for {FileName}", callingThread, Path.GetFileName(_workbookPath));
            _shutdownCts.Dispose();
        }
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

    internal static void FinalizeFailedStartupOwnedProcess(
        ExcelProcessIdentity identity,
        Func<ExcelProcessIdentity, bool> attemptTermination,
        Func<ExcelProcessIdentity, bool> confirmExit)
    {
        ArgumentNullException.ThrowIfNull(attemptTermination);
        ArgumentNullException.ThrowIfNull(confirmExit);
        if (!attemptTermination(identity) && !confirmExit(identity))
        {
            throw new InvalidOperationException(
                $"Excel process {identity.ProcessId} remained live or inaccessible after startup " +
                "cleanup. Its exact PID/start-time identity remains tracked for ProcessExit cleanup.");
        }

        SessionManager.UntrackExcelProcess(identity);
    }

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
