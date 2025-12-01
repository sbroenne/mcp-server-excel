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
    private readonly string _workbookPath; // Primary workbook path
    private readonly string[] _allWorkbookPaths; // All workbook paths (includes primary)
    private readonly bool _showExcel; // Whether to show Excel window
    private readonly ILogger<ExcelBatch> _logger;
    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private int _disposed; // 0 = not disposed, 1 = disposed (using int for Interlocked.CompareExchange)

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
    /// <param name="showExcel">Whether to show the Excel window (default: false for background automation).</param>
    public ExcelBatch(string[] workbookPaths, ILogger<ExcelBatch>? logger = null, bool showExcel = false)
    {
        if (workbookPaths == null || workbookPaths.Length == 0)
            throw new ArgumentException("At least one workbook path is required", nameof(workbookPaths));

        _allWorkbookPaths = workbookPaths;
        _workbookPath = workbookPaths[0]; // Primary workbook
        _showExcel = showExcel;
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

                // Disable macro security warnings for unattended automation
                // msoAutomationSecurityForceDisable = 3 (disable all macros, no prompts)
                // See: https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity
                tempExcel.AutomationSecurity = 3; // msoAutomationSecurityForceDisable

                // Open all workbooks in the same Excel instance
                var tempWorkbooks = new Dictionary<string, dynamic>(StringComparer.OrdinalIgnoreCase);
                dynamic? primaryWorkbook = null;

                foreach (var path in _allWorkbookPaths)
                {
                    // CRITICAL: Check if file is locked at OS level BEFORE attempting Excel COM open
                    FileAccessValidator.ValidateFileNotLocked(path);

                    // Open workbook with Excel COM
                    dynamic wb;
                    try
                    {
                        wb = tempExcel.Workbooks.Open(path);
                    }
                    catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
                    {
                        // Excel Error 1004 - File is already open or locked
                        throw FileAccessValidator.CreateFileLockedError(path, ex);
                    }

                    string normalizedPath = Path.GetFullPath(path);
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
                            catch
                            {
                                // Individual work items may fail, but keep processing queue
                                // The exception is already captured in the TaskCompletionSource
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
                    catch
                    {
                        // Unexpected error, but continue processing
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

        throw new KeyNotFoundException($"Workbook '{filePath}' is not open in this batch. " +
            $"Available workbooks: {string.Join(", ", _workbooks.Keys)}");
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

        // Wait for operation to complete (caller controls cancellation)
        try
        {
            return tcs.Task.WaitAsync(cancellationToken).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException)
        {
            _logger.LogDebug("Operation cancelled for {FileName}", Path.GetFileName(_workbookPath));
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
        _logger.LogDebug("[Thread {CallingThread}] Sleeping 100ms for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        Thread.Sleep(100);

        _logger.LogDebug("[Thread {CallingThread}] Waiting for STA thread (Id={STAThread}) to exit for {FileName}", callingThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));

        // Wait for STA thread to finish cleanup (with timeout)
        if (_staThread != null && _staThread.IsAlive)
        {
            _logger.LogDebug("[Thread {CallingThread}] Calling Join() with 10s timeout on STA={STAThread}, file={FileName}", callingThread, _staThread.ManagedThreadId, Path.GetFileName(_workbookPath));

            // Short timeout (10s) since ExcelShutdownService already handles Excel.Quit() timeout (30s)
            // If we get here and thread doesn't exit, something is badly wrong, but we've done our best
            if (!_staThread.Join(TimeSpan.FromSeconds(10)))
            {
                // STA thread didn't exit - log error but don't throw
                // The 30s quit timeout in ExcelShutdownService already tried to handle hung Excel
                _logger.LogError(
                    "[Thread {CallingThread}] STA thread (Id={STAThread}) did NOT exit within 10 seconds for {FileName}. " +
                    "This indicates Excel cleanup is severely stuck. Process will leak. " +
                    "Note: ExcelShutdownService already attempted 30s quit timeout + retries.",
                    callingThread, _staThread.ManagedThreadId, Path.GetFileName(_workbookPath));

                // Don't throw - disposal should not fail. Log the leak and continue.
                // OS will clean up when process exits.
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
