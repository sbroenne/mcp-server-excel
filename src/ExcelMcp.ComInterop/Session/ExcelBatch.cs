using System.Runtime.InteropServices;
using System.Threading.Channels;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

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
    private dynamic? _excel;
    private dynamic? _workbook; // Primary workbook
    private Dictionary<string, dynamic>? _workbooks; // All workbooks keyed by normalized path
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

                dynamic tempExcel = Activator.CreateInstance(excelType)!;
                tempExcel.Visible = _showExcel;
                tempExcel.DisplayAlerts = false;

                // Capture Excel process ID for force-kill scenarios (hung Excel, dead RPC connection)
                try
                {
                    // Excel.Application.Hwnd returns the window handle
                    // Use GetWindowThreadProcessId to get process ID directly from Hwnd
                    // This works even for hidden Excel windows (Visible=false)
                    int hwnd = tempExcel.Hwnd;
                    if (hwnd != 0)
                    {
                        uint processId = 0;
                        _ = GetWindowThreadProcessId(new IntPtr(hwnd), out processId);
                        if (processId != 0)
                        {
                            _excelProcessId = (int)processId;
                            _logger.LogDebug("Captured Excel process ID via Hwnd: {ProcessId}", _excelProcessId);
                        }
                    }

                    // Fallback: If Hwnd method failed, try finding newest EXCEL.EXE process
                    if (!_excelProcessId.HasValue)
                    {
                        var excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                            .OrderByDescending(p => p.StartTime)
                            .ToList();
                        if (excelProcesses.Count > 0)
                        {
                            _excelProcessId = excelProcesses[0].Id;
                            _logger.LogDebug("Captured Excel process ID via fallback (newest): {ProcessId}", _excelProcessId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to capture Excel process ID. Force-kill will not be available.");
                }

                // Disable macro security warnings for unattended automation
                // msoAutomationSecurityForceDisable = 3 (disable all macros, no prompts)
                // See: https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity
                tempExcel.AutomationSecurity = 3; // msoAutomationSecurityForceDisable

                // Open or create workbooks in the same Excel instance
                var tempWorkbooks = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
                dynamic? primaryWorkbook = null;

                foreach (var path in _allWorkbookPaths)
                {
                    dynamic wb;
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

                        wb = tempExcel.Workbooks.Add();

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
                        // CRITICAL: Check if file is locked at OS level BEFORE attempting Excel COM open
                        FileAccessValidator.ValidateFileNotLocked(path);

                        // Open workbook with Excel COM
                        try
                        {
                            wb = tempExcel.Workbooks.Open(path);
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

                // Message pump - process work queue until completion or cancellation
                // Use polling to avoid blocking indefinitely
                while (true)
                {
                    // Check cancellation at start of each iteration
                    if (_shutdownCts.Token.IsCancellationRequested)
                    {
                        _logger.LogDebug("Shutdown requested, exiting message pump for {FileName}", Path.GetFileName(_workbookPath));
                        break;
                    }

                    try
                    {
                        // Try to read work items, with short timeout
                        if (_workQueue.Reader.TryRead(out var work))
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
                        else
                        {
                            // No work available - check if channel is completed
                            if (_workQueue.Reader.Completion.IsCompleted)
                            {
                                _logger.LogDebug("Channel completed, exiting message pump for {FileName}", Path.GetFileName(_workbookPath));
                                break;
                            }

                            // Sleep briefly to avoid busy-waiting
                            Thread.Sleep(10);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // Shutdown requested, exit gracefully
                        _logger.LogDebug("OperationCanceledException, exiting message pump for {FileName}", Path.GetFileName(_workbookPath));
                        break;
                    }
                    catch (Exception)
                    {
                        // Unexpected error in message loop iteration — continue processing
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

                // For multi-workbook batches, close all workbooks individually before quitting Excel
                if (_workbooks != null && _workbooks.Count > 1)
                {
                    _logger.LogDebug("Closing {Count} workbooks", _workbooks.Count);
                    foreach (var kvp in _workbooks.ToList())
                    {
                        try
                        {
                            dynamic? wb = kvp.Value;
                            wb.Close(false); // Don't save - explicit save must be called
                            ComUtilities.Release(ref wb!);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to close workbook {Path}", kvp.Key);
                        }
                    }
                    _workbooks.Clear();

                    // Quit Excel after all workbooks closed
                    if (_excel != null)
                    {
                        try
                        {
                            _logger.LogDebug("Quitting Excel application");
                            _excel.Quit();
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, "Failed to quit Excel");
                        }
                        finally
                        {
                            ComUtilities.Release(ref _excel!);
                        }
                    }
                }
                else
                {
                    // Single workbook: use ExcelShutdownService for resilient shutdown
                    ExcelShutdownService.CloseAndQuit(_workbook, _excel, false, _workbookPath, _logger);
                }

                _workbook = null;
                _excel = null;
                _workbooks = null;
                _context = null;

                OleMessageFilter.Revoke();
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
            var proc = System.Diagnostics.Process.GetProcessById(_excelProcessId.Value);
            return !proc.HasExited;
        }
        catch (ArgumentException)
        {
            // Process ID doesn't exist - process has terminated
            return false;
        }
    }

    public IReadOnlyDictionary<string, dynamic> Workbooks
    {
        get
        {
            ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));
            return _workbooks ?? throw new InvalidOperationException("Workbooks not initialized");
        }
    }

    public dynamic GetWorkbook(string filePath)
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
        var writeTask = _workQueue.Writer.WriteAsync(() =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
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

        // Wait for operation to complete with timeout
        // Combine caller's cancellation token with operation timeout
        try
        {
            using var timeoutCts = new CancellationTokenSource(_operationTimeout);
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

            return tcs.Task.WaitAsync(linkedCts.Token).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            // Timeout occurred (not caller cancellation)
            _logger.LogError("Operation timed out after {Timeout} for {FileName}", _operationTimeout, Path.GetFileName(_workbookPath));
            _operationTimedOut = true; // Mark timeout for aggressive cleanup during disposal
            throw new TimeoutException(
                $"Excel operation timed out after {_operationTimeout.TotalSeconds} seconds for '{Path.GetFileName(_workbookPath)}'. " +
                "Excel may be unresponsive or the operation is taking longer than expected. " +
                "Consider increasing timeoutSeconds when opening the session.");
        }
        catch (OperationCanceledException)
        {
            _logger.LogDebug("Operation cancelled for {FileName}", Path.GetFileName(_workbookPath));
            _operationTimedOut = true; // Mark timeout for aggressive cleanup during disposal
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

        // Give the thread a moment to notice the cancellation
        _logger.LogDebug("[Thread {CallingThread}] Sleeping {DelayMs}ms for {FileName}", callingThread, ComInteropConstants.FileLockRetryDelayMs, Path.GetFileName(_workbookPath));
        Thread.Sleep(ComInteropConstants.FileLockRetryDelayMs);

        _logger.LogDebug("[Thread {CallingThread}] Waiting for STA thread (Id={STAThread}) to exit for {FileName}", callingThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));

        // Wait for STA thread to finish cleanup (with timeout)
        if (_staThread != null && _staThread.IsAlive)
        {
            // Use shorter timeout if operation timed out (Excel is likely hung)
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
                _logger.LogError(
                    "[Thread {CallingThread}] STA thread (Id={STAThread}) did NOT exit within {Timeout} for {FileName}. " +
                    "Excel cleanup is severely stuck{Reason}. Attempting force-kill.",
                    callingThread, _staThread.ManagedThreadId, joinTimeout, Path.GetFileName(_workbookPath), reasonForError);

                // Force-kill the hung Excel process
                if (_excelProcessId.HasValue)
                {
                    try
                    {
                        var excelProcess = System.Diagnostics.Process.GetProcessById(_excelProcessId.Value);
                        _logger.LogWarning(
                            "[Thread {CallingThread}] Force-killing Excel process {ProcessId} for {FileName}",
                            callingThread, _excelProcessId.Value, Path.GetFileName(_workbookPath));

                        excelProcess.Kill();
                        excelProcess.WaitForExit(5000); // Wait up to 5 seconds for process to die

                        _logger.LogInformation(
                            "[Thread {CallingThread}] Successfully force-killed Excel process {ProcessId}",
                            callingThread, _excelProcessId.Value);

                        // Now wait briefly for STA thread to exit after process killed
                        if (_staThread.Join(TimeSpan.FromSeconds(5)))
                        {
                            _logger.LogDebug("[Thread {CallingThread}] STA thread exited after force-kill", callingThread);
                        }
                        else
                        {
                            _logger.LogWarning(
                                "[Thread {CallingThread}] STA thread still stuck even after force-kill. Thread leak.",
                                callingThread);
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

        // Dispose cancellation token source
        _logger.LogDebug("[Thread {CallingThread}] Disposing CancellationTokenSource for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _shutdownCts.Dispose();

        _logger.LogDebug("[Thread {CallingThread}] Dispose COMPLETED for {FileName}", callingThread, Path.GetFileName(_workbookPath));
    }
}


