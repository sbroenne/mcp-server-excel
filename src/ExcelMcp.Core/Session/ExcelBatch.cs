using System.Runtime.InteropServices;
using System.Threading.Channels;
using Sbroenne.ExcelMcp.Core.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Implementation of IExcelBatch that manages a single Excel instance on a dedicated STA thread.
/// Ensures proper COM interop with Excel using STA apartment state and OLE message filter.
/// </summary>
internal sealed class ExcelBatch : IExcelBatch
{
    private readonly string _workbookPath;
    private readonly Channel<Func<Task>> _workQueue;
    private readonly Thread _staThread;
    private bool _disposed;

    // COM state (STA thread only)
    private dynamic? _excel;
    private dynamic? _workbook;
    private ExcelContext? _context;

    public ExcelBatch(string workbookPath)
    {
        _workbookPath = workbookPath ?? throw new ArgumentNullException(nameof(workbookPath));

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

                // Open workbook
                dynamic tempWorkbook = tempExcel.Workbooks.Open(_workbookPath);

                _excel = tempExcel;
                _workbook = tempWorkbook;
                _context = new ExcelContext(_workbookPath, _excel, _workbook);

                started.SetResult();

                // Message pump - process work queue until completion
                while (_workQueue.Reader.WaitToReadAsync().AsTask().Result)
                {
                    while (_workQueue.Reader.TryRead(out var work))
                    {
                        work().GetAwaiter().GetResult();
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
                CleanupComObjects();
                OleMessageFilter.Revoke();
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

    public async Task<T> ExecuteAsync<T>(
        Func<ExcelContext, CancellationToken, ValueTask<T>> operation,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(ExcelBatch));

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

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

        return await tcs.Task;
    }

    public Task SaveAsync(CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(ExcelBatch));

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
                string extension = Path.GetExtension(_workbookPath).ToLowerInvariant();

                // Common error codes:
                // 0x800A03EC = VBA Error 1004 (general save error, often "read-only")
                // 0x800AC472 = The file is locked for editing

                if (ex.Message.Contains("read-only") ||
                    ex.HResult == unchecked((int)0x800A03EC) ||
                    ex.HResult == unchecked((int)0x800AC472))
                {
                    try
                    {
                        // Try SaveAs as a workaround
                        int fileFormat = extension == ".xlsm" ? 52 : 51;
                        _workbook!.SaveAs(_workbookPath, fileFormat);
                        tcs.SetResult();
                    }
                    catch (Exception saveAsEx)
                    {
                        tcs.SetException(new InvalidOperationException(
                            $"Failed to save workbook '{Path.GetFileName(_workbookPath)}'. " +
                            $"File may be read-only or locked. Original error: {ex.Message}",
                            saveAsEx));
                    }
                }
                else
                {
                    tcs.SetException(new InvalidOperationException(
                        $"Failed to save workbook '{Path.GetFileName(_workbookPath)}': {ex.Message}", ex));
                }
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

        return tcs.Task;
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
            }
            catch
            {
                // Excel might already be closing, ignore to continue cleanup
            }
        }

        // Release COM objects
        void Release(object? o)
        {
            if (o != null && Marshal.IsComObject(o))
            {
                try
                {
                    Marshal.FinalReleaseComObject(o);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }
            }
        }

        Release(_workbook);
        _workbook = null;
        Release(_excel);
        _excel = null;
        _context = null;

        // Force garbage collection to release COM references
        // Two GC cycles are sufficient - one to collect, one to finalize
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        _disposed = true;

        // Complete the work queue to signal STA thread to exit
        _workQueue.Writer.Complete();

        // Wait for STA thread to finish cleanup (with timeout)
        await Task.Run(() =>
        {
            if (_staThread != null && _staThread.IsAlive)
            {
                // Give STA thread 5 seconds to cleanup gracefully
                if (!_staThread.Join(TimeSpan.FromSeconds(5)))
                {
                    // Thread didn't exit cleanly, but we can't force abort in .NET Core
                    // Log warning in production scenarios
                }
            }
        });
    }
}
