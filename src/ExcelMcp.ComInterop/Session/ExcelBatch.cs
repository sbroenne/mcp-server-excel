using System.Runtime.InteropServices;
using System.Threading.Channels;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Implementation of IExcelBatch that manages a single Excel instance on a dedicated STA thread.
/// Ensures proper COM interop with Excel using STA apartment state and OLE message filter.
/// </summary>
internal sealed class ExcelBatch : IExcelBatch
{
    private readonly string _workbookPath;
    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private readonly CancellationTokenSource _shutdownCts;
    private bool _disposed;

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

    public ExcelBatch(string workbookPath)
    {
        _workbookPath = workbookPath ?? throw new ArgumentNullException(nameof(workbookPath));
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
                while (!_shutdownCts.Token.IsCancellationRequested)
                {
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
                                break;
                            }

                            // Sleep briefly to avoid busy-waiting
                            Thread.Sleep(10);
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // Shutdown requested, exit gracefully
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
                System.Diagnostics.Debug.WriteLine($"[ExcelBatch] STA thread cleanup starting for {Path.GetFileName(_workbookPath)}");
                CleanupComObjects();
                OleMessageFilter.Revoke();
                System.Diagnostics.Debug.WriteLine($"[ExcelBatch] STA thread cleanup completed for {Path.GetFileName(_workbookPath)}");
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
        ObjectDisposedException.ThrowIf(_disposed, nameof(ExcelBatch));

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

            Console.Error.WriteLine($"[EXCEL-BATCH] TIMEOUT after {duration.TotalSeconds:F1}s (limit: {effectiveTimeout.TotalMinutes:F1}min, max: {usedMaxTimeout})");

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
        ObjectDisposedException.ThrowIf(_disposed, nameof(ExcelBatch));

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

            Console.Error.WriteLine($"[EXCEL-BATCH] TIMEOUT after {duration.TotalSeconds:F1}s (limit: {effectiveTimeout.TotalMinutes:F1}min, max: {usedMaxTimeout})");

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
        ObjectDisposedException.ThrowIf(_disposed, nameof(ExcelBatch));

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
                _excel.Quit();

                // CRITICAL: Give Excel time to fully terminate before GC
                // Excel.Quit() returns immediately but process may take time to exit
                // Without this delay, GC may run before Excel fully closes
                Thread.Sleep(2000);
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

        // CRITICAL COM cleanup pattern (Microsoft recommended):
        // Call GC.Collect() and GC.WaitForPendingFinalizers() TWICE
        // This ensures proper cleanup of Runtime Callable Wrappers (RCW)
        // Source: https://stackoverflow.com/a/38111294
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        _disposed = true;

        System.Diagnostics.Debug.WriteLine($"[ExcelBatch] DisposeAsync starting for {Path.GetFileName(_workbookPath)}");

        // Complete the work queue to signal STA thread to exit
        _workQueue.Writer.Complete();

        // Cancel the shutdown token to force message pump exit
        _shutdownCts.Cancel();

        System.Diagnostics.Debug.WriteLine($"[ExcelBatch] Waiting for STA thread to exit...");

        // Wait for STA thread to finish cleanup (with timeout)
        await Task.Run(() =>
        {
            if (_staThread != null && _staThread.IsAlive)
            {
                // Give STA thread 15 seconds to cleanup gracefully
                // (Includes time for Excel.Quit() + GC cycles)
                if (!_staThread.Join(TimeSpan.FromSeconds(15)))
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelBatch] STA thread did not exit within 5s, attempting emergency cleanup");
                    // CRITICAL: Thread didn't exit - force cleanup on this thread
                    // This is a safety measure for leaked Excel processes
                    try
                    {
                        // Attempt emergency cleanup from current thread
                        // NOTE: This violates STA rules but prevents process leaks
                        if (_excel != null)
                        {
                            try { _excel.Quit(); } catch { /* Ignore */ }
                            try { Marshal.FinalReleaseComObject(_excel); } catch { /* Ignore */ }
                        }
                        if (_workbook != null)
                        {
                            try { Marshal.FinalReleaseComObject(_workbook); } catch { /* Ignore */ }
                        }
                    }
                    catch
                    {
                        // Last resort failed, but at least we tried
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelBatch] STA thread exited successfully");
                }
            }
        });

        // Force GC to clean up any remaining COM objects
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Dispose cancellation token source
        _shutdownCts.Dispose();
    }
}
