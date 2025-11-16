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
                // Cleanup COM objects on STA thread exit (children -> parents)
                _logger.LogDebug("STA thread cleanup starting for {FileName}", Path.GetFileName(_workbookPath));

                _workbook?.Close(false); // Don't save - Save must be called explicitly
                _workbook = null;

                _excel?.Quit();
                _excel = null;

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
            try
            {
                _workbook!.Save();
                return 0;
            }
            catch (COMException ex)
            {
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

                throw new InvalidOperationException(errorMessage, ex);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Unexpected error saving workbook '{Path.GetFileName(_workbookPath)}': {ex.Message}", ex);
            }
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
            _logger.LogDebug("[Thread {CallingThread}] Calling Join() with 3s timeout on STA={STAThread}, file={FileName}", callingThread, _staThread.ManagedThreadId, Path.GetFileName(_workbookPath));

            // Give STA thread 10 seconds to cleanup gracefully
            // When multiple Excel instances are being disposed, Excel COM can deadlock
            // It's better to timeout and let Windows clean up than hang forever
            if (!_staThread.Join(TimeSpan.FromSeconds(10)))
            {
                // CRITICAL: STA thread didn't exit - this means Excel.Quit() is hung
                // Do NOT attempt cross-thread COM calls - that violates COM apartment rules
                // Instead, fail cleanly and let the OS clean up the process leak
                _logger.LogError("[Thread {CallingThread}] Join() TIMED OUT after 3s waiting for STA={STAThread}, file={FileName} - Excel COM is hung. Process will leak.", callingThread, _staThread.ManagedThreadId, Path.GetFileName(_workbookPath));

                // Throw exception to signal disposal failure
                throw new InvalidOperationException(
                    $"Excel COM cleanup timed out for '{Path.GetFileName(_workbookPath)}'. " +
                    "The STA thread did not exit within 3 seconds, indicating Excel.Quit() is hung. " +
                    "This typically occurs when Excel is showing a modal dialog or is in an unresponsive state. " +
                    "The Excel.exe process will leak and must be terminated manually.");
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
