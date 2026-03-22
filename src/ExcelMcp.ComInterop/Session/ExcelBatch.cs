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
internal sealed class ExcelBatch : IExcelBatch
{
    // P/Invoke for getting process ID from window handle
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private readonly string _workbookPath; // Primary workbook path
    private readonly string[] _allWorkbookPaths; // All workbook paths (includes primary)
    private readonly bool _showExcel; // Whether to show Excel window
    private readonly bool _createNewFile; // Whether to create a new file instead of opening existing
    private readonly bool _isMacroEnabled; // For new files: whether to create .xlsm (macro-enabled)
    private readonly TimeSpan _operationTimeout; // Timeout for individual operations
    private readonly ILogger<ExcelBatch> _logger;
    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private int _disposed; // 0 = not disposed, 1 = disposed (using int for Interlocked.CompareExchange)
    private int? _excelProcessId; // Excel.exe process ID for force-kill if needed
    private bool _operationTimedOut; // Track if an operation timed out for aggressive cleanup

    // COM state (STA thread only)
    private Excel.Application? _excel;
    private Excel.Workbook? _workbook; // Primary workbook
    private Dictionary<string, Excel.Workbook>? _workbooks; // All workbooks keyed by normalized path
    private ExcelContext? _context;

    /// <summary>
    /// Creates a new ExcelBatch for one or more workbooks.
    /// All workbooks are opened in the same Excel.Application instance, enabling cross-workbook operations.
    /// </summary>
    /// <param name="workbookPaths">Paths to Excel workbooks. First path is the primary workbook.</param>
    /// <param name="logger">Optional logger for diagnostic output. If null, uses NullLogger (no output).</param>
    /// <param name="show">Whether to show the Excel window (default: false for background automation).</param>
    /// <param name="operationTimeout">Timeout for individual operations. Default: 5 minutes.</param>
    public ExcelBatch(string[] workbookPaths, ILogger<ExcelBatch>? logger = null, bool show = false, TimeSpan? operationTimeout = null)
        : this(workbookPaths, logger, show, createNewFile: false, isMacroEnabled: false, operationTimeout: operationTimeout)
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
    /// <param name="operationTimeout">Timeout for individual operations. Default: 5 minutes.</param>
    /// <returns>ExcelBatch instance with the new workbook open.</returns>
    internal static ExcelBatch CreateNewWorkbook(string filePath, bool isMacroEnabled, ILogger<ExcelBatch>? logger = null, bool show = false, TimeSpan? operationTimeout = null)
    {
        return new ExcelBatch([filePath], logger, show, createNewFile: true, isMacroEnabled: isMacroEnabled, operationTimeout: operationTimeout);
    }

    /// <summary>
    /// Private constructor that handles both opening existing files and creating new ones.
    /// </summary>
    private ExcelBatch(string[] workbookPaths, ILogger<ExcelBatch>? logger, bool show, bool createNewFile, bool isMacroEnabled, TimeSpan? operationTimeout = null)
    {
        if (workbookPaths == null || workbookPaths.Length == 0)
            throw new ArgumentException("At least one workbook path is required", nameof(workbookPaths));

        _allWorkbookPaths = workbookPaths;
        _workbookPath = workbookPaths[0]; // Primary workbook
        _showExcel = show;
        _createNewFile = createNewFile;
        _isMacroEnabled = isMacroEnabled;
        _operationTimeout = operationTimeout ?? ComInteropConstants.DefaultOperationTimeout;
        _logger = logger ?? NullLogger<ExcelBatch>.Instance;
        _shutdownCts = new CancellationTokenSource();

        // Create unbounded channel for work items
        _workQueue = Channel.CreateUnbounded<Func<Task>>(new UnboundedChannelOptions
        {
            SingleReader = true,
            SingleWriter = false
        });

        // Start STA thread with message pump
        var started = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        _staThread = new Thread(() =>
        {
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

                Excel.Application tempExcel = (Excel.Application)Activator.CreateInstance(excelType)!;
                tempExcel.Visible = _showExcel;
                tempExcel.DisplayAlerts = false;

                // Capture Excel process ID for force-kill scenarios (hung Excel, dead RPC connection)
                // Retry with delay: Excel's HWND may not be immediately available under system load.
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
                                SessionManager.TrackExcelProcess(_excelProcessId.Value);
                                _logger.LogDebug("Captured Excel process ID via Hwnd: {ProcessId} (attempt {Attempt})",
                                    _excelProcessId, attempt);
                                break;
                            }
                        }

                        if (attempt < maxRetries)
                        {
                            _logger.LogDebug("Hwnd not available yet (attempt {Attempt}/{Max}), retrying in {Delay}ms",
                                attempt, maxRetries, retryDelayMs);
                            Thread.Sleep(retryDelayMs);
                        }
                    }

                    if (!_excelProcessId.HasValue)
                    {
                        _logger.LogWarning(
                            "Could not determine Excel process ID via Hwnd after {MaxRetries} attempts. " +
                            "Force-kill will be disabled for this session to avoid killing unrelated Excel instances.",
                            maxRetries);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to capture Excel process ID. Force-kill will not be available.");
                }

                // Disable macro security warnings for unattended automation
                // msoAutomationSecurityForceDisable = 3 (disable all macros, no prompts)
                // See: https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity
                // PIA gap: MsoAutomationSecurity is in office.dll (Microsoft.Office.Core) which is NOT bundled
                // with the Excel PIA NuGet package. Casting tempExcel to (object) first forces pure IDispatch
                // binding so the DLR never tries to load office.dll to resolve the MsoAutomationSecurity type.
                // Without (object) cast: ((dynamic)Excel.Application) retains COM type metadata → office.dll load → crash.
                ((dynamic)(object)tempExcel).AutomationSecurity = 3;

                // Open or create workbooks in the same Excel instance
                var tempWorkbooks = new Dictionary<string, Excel.Workbook>(StringComparer.OrdinalIgnoreCase);
                Excel.Workbook? primaryWorkbook = null;

                foreach (var path in _allWorkbookPaths)
                {
                    Excel.Workbook wb;
                    string normalizedPath = Path.GetFullPath(path);

                    if (_createNewFile)
                    {
                        // CREATE NEW FILE: Use Add() + SaveAs() instead of Open()
                        // Validate directory exists (do not create automatically)
                        string? directory = Path.GetDirectoryName(normalizedPath);
                        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                        {
                            throw new DirectoryNotFoundException($"Directory does not exist: '{directory}'. Create the directory first before creating Excel files.");
                        }

                        wb = (Excel.Workbook)tempExcel.Workbooks.Add();

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
                        bool isIrm = FileAccessValidator.IsIrmProtected(normalizedPath);

                        if (isIrm)
                        {
                            // IRM/AIP-protected files are OLE2 containers that cannot be opened
                            // exclusively. Excel must be visible so the user can authenticate
                            // through the IRM credential prompt.
                            tempExcel.Visible = true;
                            _logger.LogDebug(
                                "IRM-protected file detected: {FileName}. Forcing Excel visible and opening read-only.",
                                Path.GetFileName(normalizedPath));
                        }
                        else
                        {
                            // CRITICAL: Check if file is locked at OS level BEFORE attempting Excel COM open
                            FileAccessValidator.ValidateFileNotLocked(path);
                        }

                        // Open workbook with Excel COM
                        try
                        {
                            wb = isIrm
                                // ReadOnly=true prevents "exclusive access required" errors on IRM-encrypted files
                                ? (Excel.Workbook)tempExcel.Workbooks.Open(normalizedPath, ReadOnly: true)
                                : (Excel.Workbook)tempExcel.Workbooks.Open(path);
                        }
                        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
                        {
                            // Excel Error 1004 - File is already open or locked
                            throw FileAccessValidator.CreateFileLockedError(path, ex);
                        }
                    }

                    tempWorkbooks[normalizedPath] = wb;

                    if (path == _workbookPath)
                    {
                        primaryWorkbook = wb;
                    }
                }

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
                            try
                            {
                                work().GetAwaiter().GetResult();
                            }
                            catch (Exception)
                            {
                                // Individual work items may fail, but keep processing queue.
                                // The exception is already captured in the TaskCompletionSource.
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // Shutdown requested via _shutdownCts.
                        // Drain any remaining work items so in-flight Execute() callers get their
                        // results/exceptions promptly instead of waiting for the 5-minute timeout.
                        // This is safe: Excel COM objects are still alive (cleaned up in the finally
                        // block below), and Writer.Complete() prevents new items from arriving.
                        while (_workQueue.Reader.TryRead(out var remainingWork))
                        {
                            try
                            {
                                remainingWork().GetAwaiter().GetResult();
                            }
                            catch (Exception)
                            {
                                // Already captured in TaskCompletionSource
                            }
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
                }

                // Unified shutdown: use ExcelShutdownService for ALL workbook close/quit operations.
                // Previously multi-workbook batches used bare COM calls without resilience,
                // while single-workbook batches used ExcelShutdownService. Now both paths
                // get the same exponential backoff retry for COM busy conditions.
                if (_workbooks != null && _workbooks.Count > 1)
                {
                    _logger.LogDebug("Closing {Count} workbooks via ExcelShutdownService", _workbooks.Count);

                    // Close all non-primary workbooks first (without quitting Excel)
                    foreach (var kvp in _workbooks.ToList())
                    {
                        if (kvp.Value == _workbook)
                        {
                            continue; // Primary workbook closed last (with Quit)
                        }

                        // CloseAndQuit with excel=null closes workbook only, doesn't quit
                        ExcelShutdownService.CloseAndQuit(kvp.Value, null, false, kvp.Key, _logger);
                    }
                    _workbooks.Clear();

                    // Close primary workbook AND quit Excel (with resilient retry)
                    ExcelShutdownService.CloseAndQuit(_workbook, _excel, false, _workbookPath, _logger);
                }
                else
                {
                    // Single workbook: same ExcelShutdownService path
                    ExcelShutdownService.CloseAndQuit(_workbook, _excel, false, _workbookPath, _logger);
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

        // Wait for STA thread to initialize
        started.Task.GetAwaiter().GetResult();
    }

    public string WorkbookPath => _workbookPath;

    public ILogger Logger => _logger;

    public int? ExcelProcessId => _excelProcessId;

    public TimeSpan OperationTimeout => _operationTimeout;

    public bool IsExcelProcessAlive()
    {
        if (_disposed != 0) return false;
        if (!_excelProcessId.HasValue) return false;

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

        string normalizedPath = Path.GetFullPath(filePath);
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

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        // Post operation to STA thread synchronously
        // RACE CONDITION NOTE: Dispose() may call Writer.Complete() between our _disposed check
        // above and this WriteAsync() call. ChannelClosedException means the session is shutting
        // down — convert to ObjectDisposedException for a clean caller experience.
        try
        {
            var writeTask = _workQueue.Writer.WriteAsync(() =>
            {
                try
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // STRUCTURAL SAFETY: Suppress ScreenUpdating for every operation.
                    // Restores on completion or exception. Reduces COM callbacks and
                    // improves performance for bulk operations.
                    using var writeGuard = new ExcelWriteGuard((Excel.Application)_context!.App, _logger);

                    var result = operation(_context!, cancellationToken);
                    tcs.SetResult(result);
                }
                catch (OperationCanceledException oce)
                {
                    tcs.TrySetCanceled(oce.CancellationToken);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
                return Task.CompletedTask;
            }, cancellationToken);

            // ValueTask is completed synchronously in normal case
            if (writeTask.IsCompleted)
            {
                writeTask.GetAwaiter().GetResult();
            }
            else
            {
                // Fallback: should not normally occur with unbounded channel
                writeTask.AsTask().GetAwaiter().GetResult();
            }
        }
        catch (ChannelClosedException)
        {
            // Dispose() completed the channel between our _disposed check and WriteAsync.
            // The session is shutting down — report as disposed.
            throw new ObjectDisposedException(nameof(ExcelBatch),
                $"Session for '{Path.GetFileName(_workbookPath)}' was disposed while submitting an operation.");
        }

        // Wait for operation to complete with timeout.
        // When the caller provides a cancellation token (e.g., PowerQuery refresh with its own timeout),
        // respect it exclusively and don't layer the session _operationTimeout on top.
        // This prevents a double-cap where min(callerTimeout, sessionTimeout) is always the shorter one —
        // which caused heavy Power Query refreshes (~8+ min) to always fail against the 5-min default.
        try
        {
            if (cancellationToken.CanBeCanceled)
            {
                // Caller controls the timeout — use their token exclusively
                return tcs.Task.WaitAsync(cancellationToken).GetAwaiter().GetResult();
            }
            else
            {
                // No caller timeout — apply session-level operation timeout as safety net
                using var timeoutCts = new CancellationTokenSource(_operationTimeout);
                return tcs.Task.WaitAsync(timeoutCts.Token).GetAwaiter().GetResult();
            }
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            // Session timeout occurred (not caller cancellation) — only happens in the else branch
            _logger.LogError("Operation timed out after {Timeout} for {FileName}", _operationTimeout, Path.GetFileName(_workbookPath));
            _operationTimedOut = true; // Mark timeout for aggressive cleanup during disposal
            throw new TimeoutException(
                $"Excel operation timed out after {_operationTimeout.TotalSeconds} seconds for '{Path.GetFileName(_workbookPath)}'. " +
                "Excel may be unresponsive or the operation is taking longer than expected. " +
                "Consider increasing timeoutSeconds when opening the session.");
        }
        catch (OperationCanceledException)
        {
            _logger.LogDebug("Operation cancelled or timed out for {FileName}", Path.GetFileName(_workbookPath));
            _operationTimedOut = true; // STA thread may still be blocked — session is unusable
            throw;
        }
    }

    public void Save(CancellationToken cancellationToken = default)
    {
        Execute((ctx, ct) =>
        {
            ExcelShutdownService.SaveWorkbookWithTimeout(
                _workbook!,
                Path.GetFileName(_workbookPath),
                _logger,
                ct);
            return 0;
        }, cancellationToken);
    }

    public void Dispose()
    {
        var callingThread = Environment.CurrentManagedThreadId;

        // Use Interlocked.CompareExchange for thread-safe disposal check
        // Returns 0 if exchange succeeded (was not disposed), 1 if already disposed
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
        {
            _logger.LogDebug("[Thread {CallingThread}] Dispose skipped - already disposed for {FileName}", callingThread, Path.GetFileName(_workbookPath));
            return; // Already disposed
        }

        _logger.LogDebug("[Thread {CallingThread}] Dispose starting for {FileName}", callingThread, Path.GetFileName(_workbookPath));

        // Cancel the shutdown token FIRST to wake up the message pump
        _logger.LogDebug("[Thread {CallingThread}] Cancelling shutdown token for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _shutdownCts.Cancel();

        // Then complete the work queue
        _logger.LogDebug("[Thread {CallingThread}] Completing work queue for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _workQueue.Writer.Complete();

        _logger.LogDebug("[Thread {CallingThread}] Waiting for STA thread (Id={STAThread}) to exit for {FileName}", callingThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));

        // When operation timed out, the STA thread is stuck in IDispatch.Invoke (unmanaged COM call
        // that cannot be cancelled). Kill the Excel process FIRST to unblock the STA thread, then wait.
        if (_operationTimedOut && _excelProcessId.HasValue && _staThread != null && _staThread.IsAlive)
        {
            // INSTRUMENTATION: Track pre-emptive kill path entry
            _logger.LogWarning(
                "[DIAG-DISPOSE-TIMEOUT-PREKILL] [Thread {CallingThread}] Operation timed out — force-killing Excel process {ProcessId} BEFORE waiting for STA thread to unblock IDispatch.Invoke for {FileName}",
                callingThread, _excelProcessId.Value, Path.GetFileName(_workbookPath));
            SessionDiagnostics.WriteStdErr(
                $"[DIAG-DISPOSE-TIMEOUT-PREKILL] [Thread {callingThread}] Operation timed out — force-killing Excel process {_excelProcessId.Value} BEFORE waiting for STA thread to unblock IDispatch.Invoke for {Path.GetFileName(_workbookPath)}");
            try
            {
                using var excelProcess = System.Diagnostics.Process.GetProcessById(_excelProcessId.Value);
                if (!excelProcess.HasExited)
                {
                    excelProcess.Kill();
                    excelProcess.WaitForExit(5000);
                    // INSTRUMENTATION: Track successful pre-emptive kill
                    _logger.LogInformation(
                        "[DIAG-DISPOSE-TIMEOUT-PREKILL-SUCCESS] [Thread {CallingThread}] Force-killed Excel process {ProcessId} (pre-emptive, before STA join)",
                        callingThread, _excelProcessId.Value);
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-TIMEOUT-PREKILL-SUCCESS] [Thread {callingThread}] Force-killed Excel process {_excelProcessId.Value} (pre-emptive, before STA join)");
                }
                else
                {
                    // INSTRUMENTATION: Process already gone
                    _logger.LogDebug(
                        "[DIAG-DISPOSE-TIMEOUT-PREKILL-ALREADY-GONE] [Thread {CallingThread}] Excel process {ProcessId} already exited",
                        callingThread, _excelProcessId.Value);
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-TIMEOUT-PREKILL-ALREADY-GONE] [Thread {callingThread}] Excel process {_excelProcessId.Value} already exited");
                }
            }
            catch (ArgumentException)
            {
                _logger.LogDebug("[DIAG-DISPOSE-TIMEOUT-PREKILL-NOT-FOUND] [Thread {CallingThread}] Excel process {ProcessId} not found (already exited)", callingThread, _excelProcessId.Value);
                SessionDiagnostics.WriteStdErr(
                    $"[DIAG-DISPOSE-TIMEOUT-PREKILL-NOT-FOUND] [Thread {callingThread}] Excel process {_excelProcessId.Value} not found (already exited)");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[DIAG-DISPOSE-TIMEOUT-PREKILL-FAILED] [Thread {CallingThread}] Failed to force-kill Excel process {ProcessId}", callingThread, _excelProcessId.Value);
                SessionDiagnostics.WriteStdErr(
                    $"[DIAG-DISPOSE-TIMEOUT-PREKILL-FAILED] [Thread {callingThread}] Failed to force-kill Excel process {_excelProcessId.Value}: {ex.GetType().Name}: {ex.Message}");
            }
        }

        // Wait for STA thread to finish cleanup (with timeout)
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

                // Force-kill the hung Excel process
                if (_excelProcessId.HasValue)
                {
                    try
                    {
                        using var excelProcess = System.Diagnostics.Process.GetProcessById(_excelProcessId.Value);
                        // INSTRUMENTATION: Track force-kill attempt after STA join timeout
                        _logger.LogWarning(
                            "[DIAG-DISPOSE-FORCE-KILL-ATTEMPT] [Thread {CallingThread}] Force-killing Excel process {ProcessId} for {FileName}",
                            callingThread, _excelProcessId.Value, Path.GetFileName(_workbookPath));
                        SessionDiagnostics.WriteStdErr(
                            $"[DIAG-DISPOSE-FORCE-KILL-ATTEMPT] [Thread {callingThread}] Force-killing Excel process {_excelProcessId.Value} for {Path.GetFileName(_workbookPath)}");

                        excelProcess.Kill();
                        excelProcess.WaitForExit(5000); // Wait up to 5 seconds for process to die

                        // INSTRUMENTATION: Track successful force-kill
                        _logger.LogInformation(
                            "[DIAG-DISPOSE-FORCE-KILL-SUCCESS] [Thread {CallingThread}] Successfully force-killed Excel process {ProcessId}",
                            callingThread, _excelProcessId.Value);
                        SessionDiagnostics.WriteStdErr(
                            $"[DIAG-DISPOSE-FORCE-KILL-SUCCESS] [Thread {callingThread}] Successfully force-killed Excel process {_excelProcessId.Value}");

                        // Now wait briefly for STA thread to exit after process killed
                        if (_staThread.Join(TimeSpan.FromSeconds(5)))
                        {
                            // INSTRUMENTATION: Track STA thread exit after force-kill
                            _logger.LogDebug("[DIAG-DISPOSE-STA-EXIT-AFTER-KILL] [Thread {CallingThread}] STA thread exited after force-kill", callingThread);
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-DISPOSE-STA-EXIT-AFTER-KILL] [Thread {callingThread}] STA thread exited after force-kill");
                        }
                        else
                        {
                            // INSTRUMENTATION: Track persistent STA thread leak
                            _logger.LogWarning(
                                "[DIAG-DISPOSE-STA-LEAK] [Thread {CallingThread}] STA thread still stuck even after force-kill. Thread leak.",
                                callingThread);
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-DISPOSE-STA-LEAK] [Thread {callingThread}] STA thread still stuck even after force-kill. Thread leak.");
                        }
                    }
                    catch (ArgumentException)
                    {
                        _logger.LogWarning(
                            "[Thread {CallingThread}] Excel process {ProcessId} not found (already exited?)",
                            callingThread, _excelProcessId.Value);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex,
                            "[Thread {CallingThread}] Failed to force-kill Excel process {ProcessId}",
                            callingThread, _excelProcessId.Value);
                    }
                }
                else
                {
                    _logger.LogError(
                        "[Thread {CallingThread}] No Excel process ID captured - cannot force-kill. Process will leak.",
                        callingThread);
                }
            }
        }
        else
        {
            _logger.LogDebug("[Thread {CallingThread}] STA thread was null or not alive for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        }

        // Wait for Excel process to fully terminate to prevent CO_E_SERVER_EXEC_FAILURE
        // on subsequent Activator.CreateInstance calls. excel.Quit() + COM release doesn't
        // guarantee the EXCEL.EXE process has exited — rapid create/destroy cycles can fail.
        if (_excelProcessId.HasValue)
        {
            try
            {
                using var excelProc = System.Diagnostics.Process.GetProcessById(_excelProcessId.Value);
                if (!excelProc.HasExited)
                {
                    // INSTRUMENTATION: Track process linger detection
                    _logger.LogDebug(
                        "[DIAG-DISPOSE-PROCESS-WAIT] [Thread {CallingThread}] Waiting for Excel process {ProcessId} to exit for {FileName}",
                        callingThread, _excelProcessId.Value, Path.GetFileName(_workbookPath));
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-PROCESS-WAIT] [Thread {callingThread}] Waiting for Excel process {_excelProcessId.Value} to exit for {Path.GetFileName(_workbookPath)}");

                    if (!excelProc.WaitForExit(5000))
                    {
                        // INSTRUMENTATION: Track process linger timeout (THE KEY SYMPTOM)
                        _logger.LogWarning(
                            "[DIAG-DISPOSE-PROCESS-LINGER] [Thread {CallingThread}] Excel process {ProcessId} did not exit within 5s for {FileName}. Force-killing to prevent zombie accumulation.",
                            callingThread, _excelProcessId.Value, Path.GetFileName(_workbookPath));
                        SessionDiagnostics.WriteStdErr(
                            $"[DIAG-DISPOSE-PROCESS-LINGER] [Thread {callingThread}] Excel process {_excelProcessId.Value} did not exit within 5s for {Path.GetFileName(_workbookPath)}. Force-killing to prevent zombie accumulation.");

                        // Force-kill: Excel was already told to Quit() and COM refs were released.
                        // A process still running after 5s is hung and will leak desktop resources.
                        try
                        {
                            excelProc.Kill();
                            excelProc.WaitForExit(3000);
                            // INSTRUMENTATION: Track final force-kill of lingering process
                            _logger.LogInformation(
                                "[DIAG-DISPOSE-PROCESS-LINGER-KILLED] [Thread {CallingThread}] Force-killed lingering Excel process {ProcessId} for {FileName}",
                                callingThread, _excelProcessId.Value, Path.GetFileName(_workbookPath));
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-DISPOSE-PROCESS-LINGER-KILLED] [Thread {callingThread}] Force-killed lingering Excel process {_excelProcessId.Value} for {Path.GetFileName(_workbookPath)}");
                        }
                        catch (Exception killEx)
                        {
                            // INSTRUMENTATION: Track final force-kill failure (CRITICAL - this is the leak)
                            _logger.LogWarning(killEx,
                                "[DIAG-DISPOSE-PROCESS-LINGER-KILL-FAILED] [Thread {CallingThread}] Failed to force-kill Excel process {ProcessId}",
                                callingThread, _excelProcessId.Value);
                            SessionDiagnostics.WriteStdErr(
                                $"[DIAG-DISPOSE-PROCESS-LINGER-KILL-FAILED] [Thread {callingThread}] Failed to force-kill Excel process {_excelProcessId.Value}: {killEx.GetType().Name}: {killEx.Message}");
                        }
                    }
                    else
                    {
                        // INSTRUMENTATION: Track normal process exit
                        _logger.LogDebug(
                            "[DIAG-DISPOSE-PROCESS-EXITED] [Thread {CallingThread}] Excel process {ProcessId} exited normally for {FileName}",
                            callingThread, _excelProcessId.Value, Path.GetFileName(_workbookPath));
                        SessionDiagnostics.WriteStdErr(
                            $"[DIAG-DISPOSE-PROCESS-EXITED] [Thread {callingThread}] Excel process {_excelProcessId.Value} exited normally for {Path.GetFileName(_workbookPath)}");
                    }
                }
                else
                {
                    // INSTRUMENTATION: Process already gone (fast path)
                    _logger.LogDebug(
                        "[DIAG-DISPOSE-PROCESS-ALREADY-GONE] [Thread {CallingThread}] Excel process {ProcessId} already exited for {FileName}",
                        callingThread, _excelProcessId.Value, Path.GetFileName(_workbookPath));
                    SessionDiagnostics.WriteStdErr(
                        $"[DIAG-DISPOSE-PROCESS-ALREADY-GONE] [Thread {callingThread}] Excel process {_excelProcessId.Value} already exited for {Path.GetFileName(_workbookPath)}");
                }
            }
            catch (ArgumentException)
            {
                // Process already terminated — this is the expected fast path
                _logger.LogDebug(
                    "[DIAG-DISPOSE-PROCESS-NOT-FOUND] [Thread {CallingThread}] Excel process {ProcessId} not found (already exited)",
                    callingThread, _excelProcessId.Value);
                SessionDiagnostics.WriteStdErr(
                    $"[DIAG-DISPOSE-PROCESS-NOT-FOUND] [Thread {callingThread}] Excel process {_excelProcessId.Value} not found (already exited)");
            }
            catch (InvalidOperationException ex)
            {
                // Process object is not associated with a running process
                _logger.LogDebug(ex,
                    "[DIAG-DISPOSE-PROCESS-INVALID] [Thread {CallingThread}] Excel process {ProcessId} object invalid",
                    callingThread, _excelProcessId.Value);
                SessionDiagnostics.WriteStdErr(
                    $"[DIAG-DISPOSE-PROCESS-INVALID] [Thread {callingThread}] Excel process {_excelProcessId.Value} object invalid: {ex.Message}");
            }
        }

        // Dispose cancellation token source
        _logger.LogDebug("[Thread {CallingThread}] Disposing CancellationTokenSource for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _shutdownCts.Dispose();

        _logger.LogDebug("[Thread {CallingThread}] Dispose COMPLETED for {FileName}", callingThread, Path.GetFileName(_workbookPath));
    }

}

