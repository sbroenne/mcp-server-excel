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
    private readonly string _workbookPath;
    private readonly ILogger<ExcelBatch> _logger;
    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private int _disposed; // 0 = not disposed, 1 = disposed (using int for Interlocked.CompareExchange)

    /// <summary>
    /// Default timeout for most Excel operations (list, get, set, etc.).
    /// </summary>
    private static readonly TimeSpan DefaultOperationTimeout = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Maximum allowed timeout to prevent runaway operations.
    /// Heavy operations like refresh can request up to this limit.
    /// </summary>
    private static readonly TimeSpan MaxOperationTimeout = TimeSpan.FromMinutes(5);

    // COM state (STA thread only)
    private dynamic? _excel;
    private dynamic? _workbook;
    private ExcelContext? _context;

    /// <summary>
    /// Creates a new ExcelBatch for the specified workbook.
    /// </summary>
    /// <param name="workbookPath">Path to the Excel workbook</param>
    /// <param name="logger">Optional logger for diagnostic output. If null, uses NullLogger (no output).</param>
    public ExcelBatch(string workbookPath, ILogger<ExcelBatch>? logger = null)
    {
        _workbookPath = workbookPath ?? throw new ArgumentNullException(nameof(workbookPath));
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
                tempExcel.Visible = false;
                tempExcel.DisplayAlerts = false;

                // Disable macro security warnings for unattended automation
                // msoAutomationSecurityForceDisable = 3 (disable all macros, no prompts)
                // See: https://learn.microsoft.com/en-us/office/vba/api/word.application.automationsecurity
                tempExcel.AutomationSecurity = 3; // msoAutomationSecurityForceDisable

                // CRITICAL: Check if file is locked at OS level BEFORE attempting Excel COM open
                // This fails fast without the overhead of Excel COM initialization
                FileAccessValidator.ValidateFileNotLocked(_workbookPath);

                // Open workbook with Excel COM
                dynamic tempWorkbook;
                try
                {
                    tempWorkbook = tempExcel.Workbooks.Open(_workbookPath);
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
                {
                    // Excel Error 1004 - File is already open or locked
                    // This is a backup catch in case OS-level check missed something
                    throw FileAccessValidator.CreateFileLockedError(_workbookPath, ex);
                }

                _excel = tempExcel;
                _workbook = tempWorkbook;
                _context = new ExcelContext(_workbookPath, _excel, _workbook);

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
                // Cleanup on STA thread exit
                _logger.LogDebug("STA thread cleanup starting for {FileName}", Path.GetFileName(_workbookPath));
                CleanupComObjects();
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

    // Synchronous COM operations
    public async Task<T> Execute<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));

        // Clamp timeout between default and max
        var effectiveTimeout = timeout.HasValue
            ? TimeSpan.FromMilliseconds(Math.Min(timeout.Value.TotalMilliseconds, MaxOperationTimeout.TotalMilliseconds))
            : DefaultOperationTimeout;

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var startTime = DateTime.UtcNow;

        // Post operation to STA thread
        await _workQueue.Writer.WriteAsync(() =>
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

        // Add timeout protection
        using var timeoutCts = new CancellationTokenSource(effectiveTimeout);
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

        try
        {
            return await tcs.Task.WaitAsync(linkedCts.Token);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
        {
            var duration = DateTime.UtcNow - startTime;
            var usedMaxTimeout = effectiveTimeout >= MaxOperationTimeout;

            _logger.LogWarning("Operation timeout after {Duration}s (limit: {Limit}min, max: {UsedMax}) for {FileName}",
                duration.TotalSeconds, effectiveTimeout.TotalMinutes, usedMaxTimeout, Path.GetFileName(_workbookPath));

            var message = usedMaxTimeout
                ? $"Excel operation exceeded maximum timeout of {MaxOperationTimeout.TotalMinutes} minutes (actual: {duration.TotalMinutes:F1} min). " +
                  "This indicates Excel is hung, unresponsive, or the operation is too complex. " +
                  "Check if Excel is showing a dialog or consider breaking the operation into smaller steps."
                : $"Excel operation timed out after {effectiveTimeout.TotalMinutes} minutes (actual: {duration.TotalMinutes:F1} min). " +
                  $"For large datasets or complex operations, more time may be needed (maximum: {MaxOperationTimeout.TotalMinutes} min).";

            throw new TimeoutException(message);
        }
    }

    // Genuinely async operations (file I/O, etc.)
    public async Task<T> ExecuteAsync<T>(
        Func<ExcelContext, CancellationToken, Task<T>> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));

        // Clamp timeout between default and max
        var effectiveTimeout = timeout.HasValue
            ? TimeSpan.FromMilliseconds(Math.Min(timeout.Value.TotalMilliseconds, MaxOperationTimeout.TotalMilliseconds))
            : DefaultOperationTimeout;

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var startTime = DateTime.UtcNow;

        // Post operation to STA thread
        await _workQueue.Writer.WriteAsync(async () =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                var result = await operation(_context!, cancellationToken);
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
        }, cancellationToken);

        // Add timeout protection
        using var timeoutCts = new CancellationTokenSource(effectiveTimeout);
        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

        try
        {
            return await tcs.Task.WaitAsync(linkedCts.Token);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
        {
            var duration = DateTime.UtcNow - startTime;
            var usedMaxTimeout = effectiveTimeout >= MaxOperationTimeout;

            _logger.LogWarning(
                "Async operation timeout after {Duration}s (limit: {Limit}min, max: {UsedMax}) for {FileName}",
                duration.TotalSeconds, effectiveTimeout.TotalMinutes, usedMaxTimeout, Path.GetFileName(_workbookPath));

            var message = usedMaxTimeout
                ? $"Excel operation exceeded maximum timeout of {MaxOperationTimeout.TotalMinutes} minutes (actual: {duration.TotalMinutes:F1} min). " +
                  "This indicates Excel is hung, unresponsive, or the operation is too complex. " +
                  "Check if Excel is showing a dialog or consider breaking the operation into smaller steps."
                : $"Excel operation timed out after {effectiveTimeout.TotalMinutes} minutes (actual: {duration.TotalMinutes:F1} min). " +
                  $"For large datasets or complex operations, more time may be needed (maximum: {MaxOperationTimeout.TotalMinutes} min).";

            throw new TimeoutException(message);
        }
    }

    public async Task SaveAsync(TimeSpan? timeout = null, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed != 0, nameof(ExcelBatch));

        // Determine effective timeout (save operations default 2 minutes, no maximum limit)
        var effectiveTimeout = timeout ?? TimeSpan.FromMinutes(2); // Save operations get 2-minute default

        var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        // Post save operation to STA thread
        _workQueue.Writer.TryWrite(async () =>
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                _workbook!.Save();
                tcs.SetResult();
            }
            catch (COMException ex)
            {
                // Map common Excel COM error codes to meaningful messages
                string errorMessage = ex.HResult switch
                {
                    unchecked((int)0x800A03EC) =>
                        $"Cannot save '{Path.GetFileName(_workbookPath)}'. " +
                        "The file may be read-only, locked by another process, or the path may not exist.",
                    unchecked((int)0x800AC472) =>
                        $"Cannot save '{Path.GetFileName(_workbookPath)}'. " +
                        "The file is locked for editing by another user or process.",
                    _ => $"Failed to save workbook '{Path.GetFileName(_workbookPath)}': {ex.Message}"
                };

                tcs.SetException(new InvalidOperationException(errorMessage, ex));
            }
            catch (OperationCanceledException oce)
            {
                tcs.TrySetCanceled(oce.CancellationToken);
            }
            catch (Exception ex)
            {
                tcs.SetException(new InvalidOperationException(
                    $"Unexpected error saving workbook '{Path.GetFileName(_workbookPath)}': {ex.Message}", ex));
            }

            await Task.CompletedTask;
        });

        // Apply timeout protection
        try
        {
            await tcs.Task.WaitAsync(effectiveTimeout, cancellationToken);
        }
        catch (TimeoutException)
        {
            // Rethrow with contextual message
            throw new TimeoutException(
                $"Excel save operation timed out after {effectiveTimeout.TotalMinutes:F2} minutes. " +
                $"The workbook '{Path.GetFileName(_workbookPath)}' may have many changes or Excel may be busy. " +
                $"Maximum timeout is {MaxOperationTimeout.TotalMinutes} minutes.");
        }
    }

    private static bool IsTransientSaveError(COMException ex)
    {
        // Excel temp file access errors during concurrent saves (transient)
        // "Microsoft Excel cannot access the file 'C:\...\Temp\...'..."
        if (ex.Message.Contains("Temp", StringComparison.OrdinalIgnoreCase) &&
            (ex.Message.Contains("cannot access") || ex.Message.Contains("being used by another program")))
        {
            return true;
        }

        // Common transient error codes
        return ex.HResult switch
        {
            unchecked((int)0x800A03EC) => true,  // VBA Error 1004 (can be transient during concurrent ops)
            unchecked((int)0x800AC472) => true,  // File locked (can be transient)
            _ => false
        };
    }

    /// <summary>
    /// Cleanup COM objects in reverse order (children -> parents).
    /// MUST be called on STA thread.
    /// </summary>
    private void CleanupComObjects()
    {
        // Close workbook without saving (SaveAsync must be called explicitly)
        if (_workbook != null)
        {
            try
            {
                _workbook.Close(false);
            }
            catch
            {
                // Workbook might already be closed, ignore to continue cleanup
            }
        }

        // Quit Excel application
        if (_excel != null)
        {
            try
            {
                // Excel.Quit() can deadlock when multiple instances are disposing
                // Use a timeout task to prevent hanging forever
                var quitTask = Task.Run(() =>
                {
                    try
                    {
                        _excel.Quit();
                    }
                    catch
                    {
                        // Ignore - Excel might already be closing
                    }
                });

                // Wait up to 2 seconds for Quit() to complete
                if (!quitTask.Wait(TimeSpan.FromSeconds(2)))
                {
                    // Quit() is hung - likely Excel COM deadlock
                    // Let it hang in background and continue cleanup
                    // Windows will eventually clean up the process
                }

                // Note: Excel.Quit() returns immediately but process may take time to exit
                // We rely on DisposeAsync's thread join to wait for actual termination
            }
            catch
            {
                // Excel might already be closing, ignore to continue cleanup
            }
        }

        // Null out references to allow GC to clean up
        // DON'T use Marshal.ReleaseComObject - let GC handle it
        // Microsoft recommends against using ReleaseComObject in most cases
        _workbook = null;
        _excel = null;
        _context = null;

        // Force GC to clean up COM objects (Runtime Callable Wrappers)
        // Single cycle is sufficient for modern .NET
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    public async ValueTask DisposeAsync()
    {
        var callingThread = Environment.CurrentManagedThreadId;

        // Use Interlocked.CompareExchange for thread-safe disposal check
        // Returns 0 if exchange succeeded (was not disposed), 1 if already disposed
        if (Interlocked.CompareExchange(ref _disposed, 1, 0) != 0)
        {
            _logger.LogDebug("[Thread {CallingThread}] DisposeAsync skipped - already disposed for {FileName}", callingThread, Path.GetFileName(_workbookPath));
            return; // Already disposed
        }

        _logger.LogDebug("[Thread {CallingThread}] DisposeAsync starting for {FileName}", callingThread, Path.GetFileName(_workbookPath));

        // Cancel the shutdown token FIRST to wake up the message pump
        _logger.LogDebug("[Thread {CallingThread}] Cancelling shutdown token for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _shutdownCts.Cancel();

        // Then complete the work queue
        _logger.LogDebug("[Thread {CallingThread}] Completing work queue for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _workQueue.Writer.Complete();

        // Give the thread a moment to notice the cancellation
        _logger.LogDebug("[Thread {CallingThread}] Delaying 100ms for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        await Task.Delay(100);

        _logger.LogDebug("[Thread {CallingThread}] Waiting for STA thread (Id={STAThread}) to exit for {FileName}", callingThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));

        // Wait for STA thread to finish cleanup (with timeout)
        // Use Task.Run with timeout to prevent hanging forever if Excel COM deadlocks
        var joinTask = Task.Run(() =>
        {
            var workerThread = Environment.CurrentManagedThreadId;
            _logger.LogDebug("[Thread {WorkerThread}] Task.Run worker thread starting join wait for STA={STAThread}, file={FileName}", workerThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));

            if (_staThread != null && _staThread.IsAlive)
            {
                _logger.LogDebug("[Thread {WorkerThread}] Calling Join() with 3s timeout on STA={STAThread}, file={FileName}", workerThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));

                // Give STA thread 3 seconds to cleanup gracefully
                // When multiple Excel instances are being disposed, Excel COM can deadlock
                // It's better to timeout and let Windows clean up than hang forever
                if (!_staThread.Join(TimeSpan.FromSeconds(3)))
                {
                    _logger.LogWarning("[Thread {WorkerThread}] Join() TIMED OUT after 3s waiting for STA={STAThread}, file={FileName} - attempting emergency cleanup", workerThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));
                    // CRITICAL: Thread didn't exit - force cleanup on this thread
                    // This is a safety measure for leaked Excel processes
                    try
                    {
                        // Attempt emergency cleanup from current thread
                        // NOTE: This violates STA rules but prevents process leaks
                        if (_excel != null)
                        {
                            _logger.LogDebug("[Thread {WorkerThread}] Emergency: Calling _excel.Quit() for {FileName}", workerThread, Path.GetFileName(_workbookPath));
                            try { _excel.Quit(); } catch { /* Ignore */ }
                            _logger.LogDebug("[Thread {WorkerThread}] Emergency: Releasing _excel COM object for {FileName}", workerThread, Path.GetFileName(_workbookPath));
                            try { Marshal.FinalReleaseComObject(_excel); } catch { /* Ignore */ }
                        }
                        if (_workbook != null)
                        {
                            _logger.LogDebug("[Thread {WorkerThread}] Emergency: Releasing _workbook COM object for {FileName}", workerThread, Path.GetFileName(_workbookPath));
                            try { Marshal.FinalReleaseComObject(_workbook); } catch { /* Ignore */ }
                        }
                    }
                    catch (Exception emergEx)
                    {
                        _logger.LogWarning(emergEx, "[Thread {WorkerThread}] Emergency cleanup failed for {FileName}", workerThread, Path.GetFileName(_workbookPath));
                    }
                }
                else
                {
                    _logger.LogDebug("[Thread {WorkerThread}] Join() SUCCEEDED - STA thread (Id={STAThread}) exited successfully for {FileName}", workerThread, _staThread?.ManagedThreadId ?? -1, Path.GetFileName(_workbookPath));
                }
            }
            else
            {
                _logger.LogDebug("[Thread {WorkerThread}] STA thread was null or not alive for {FileName}", workerThread, Path.GetFileName(_workbookPath));
            }
        });

        // Wait for join task with aggressive timeout (Excel.Quit can block forever in COM deadlock scenarios)
        try
        {
            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
            await joinTask.WaitAsync(timeoutCts.Token);
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("[Thread {CallingThread}] Join task timed out after 2s for {FileName} - Excel COM deadlock, force-killing Excel process", callingThread, Path.GetFileName(_workbookPath));

            // CRITICAL: Excel COM is deadlocked - force-kill all Excel processes to prevent leaks
            // This is a last resort when Excel.Quit() hangs due to COM threading issues
            try
            {
                var excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (var process in excelProcesses)
                {
                    try
                    {
                        _logger.LogWarning("[Thread {CallingThread}] Force-killing Excel process {ProcessId}", callingThread, process.Id);
                        process.Kill();
                        process.WaitForExit(1000); // Wait up to 1s for termination
                    }
                    catch (Exception killEx)
                    {
                        _logger.LogWarning(killEx, "[Thread {CallingThread}] Failed to kill Excel process {ProcessId}", callingThread, process.Id);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[Thread {CallingThread}] Failed to enumerate/kill Excel processes", callingThread);
            }
        }

        // Force GC to clean up any remaining COM objects
        _logger.LogDebug("[Thread {CallingThread}] Running GC.Collect() for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Dispose cancellation token source
        _logger.LogDebug("[Thread {CallingThread}] Disposing CancellationTokenSource for {FileName}", callingThread, Path.GetFileName(_workbookPath));
        _shutdownCts.Dispose();

        _logger.LogDebug("[Thread {CallingThread}] DisposeAsync COMPLETED for {FileName}", callingThread, Path.GetFileName(_workbookPath));
    }
}
