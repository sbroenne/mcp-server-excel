using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Suppresses screen updating unless an open workbook has a Data Model, and restores changes on disposal.
/// Events and calculation are unchanged; individual commands own any suppression.
///
/// This guard is integrated into <see cref="Session.ExcelBatch.Execute{T}"/>.
///
/// Reentrant-safe: nested guards are no-ops (ref-counted via thread-static counter).
/// </summary>
public sealed class ExcelWriteGuard : IDisposable
{
    // Thread-static ref count: only the outermost guard captures/restores state.
    // Thread-static is correct because Excel COM operations run on a dedicated STA thread.
    [ThreadStatic]
    private static int _nestingDepth;

    private readonly Excel.Application? _app;
    private readonly ILogger _logger;
    private readonly bool _isOutermost;

    private bool _screenUpdatingSuppressed;

    private bool _disposed;

    /// <summary>
    /// Creates a guard that suppresses screen updating unless an open workbook has a Data Model.
    /// Only the outermost guard in a nested chain captures and restores state.
    /// </summary>
    /// <param name="app">Excel Application COM object</param>
    /// <param name="logger">Optional logger for diagnostics</param>
    public ExcelWriteGuard(Excel.Application app, ILogger? logger = null)
    {
        _app = app;
        _logger = logger ?? NullLogger.Instance;

        _nestingDepth++;
        _isOutermost = _nestingDepth == 1;

        if (!_isOutermost)
        {
            return;
        }

        try
        {
            // Toggling this flag can dirty a saved Data Model without any workbook edits.
            // It affects every open workbook in the application, not just the selected one.
            if (_app.ScreenUpdating && !HasDataModel(_app))
            {
                _app.ScreenUpdating = false;
                _screenUpdatingSuppressed = true;
            }
        }
        catch
        {
            // Construction failed, so Dispose cannot balance the nesting count.
            _nestingDepth--;
            throw;
        }
    }

    private static bool HasDataModel(Excel.Application app)
    {
        Excel.Workbooks? books = null;
        try
        {
            books = app.Workbooks;
            int bookCount = books.Count;
            for (int bookIndex = 1; bookIndex <= bookCount; bookIndex++)
            {
                Excel.Workbook? book = null;
                Excel.Connections? connections = null;
                try
                {
                    book = books[bookIndex];
                    connections = book.Connections;
                    int connectionCount = connections.Count;
                    for (int connectionIndex = 1; connectionIndex <= connectionCount; connectionIndex++)
                    {
                        Excel.WorkbookConnection? connection = null;
                        try
                        {
                            connection = connections.Item(connectionIndex);
                            if (connection.Type == Excel.XlConnectionType.xlConnectionTypeMODEL)
                                return true;
                        }
                        finally
                        {
                            ComUtilities.Release(ref connection);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connections);
                    ComUtilities.Release(ref book);
                }
            }
            return false;
        }
        finally
        {
            ComUtilities.Release(ref books);
        }
    }

    /// <summary>
    /// Restores the original screen-updating state.
    /// Only the outermost guard restores — inner guards are no-ops.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "ScreenUpdating restoration is cleanup after the operation; log its exact COM failure without replacing the primary operation exception. This does not suppress operation reads or writes.")]
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _nestingDepth--;

        if (!_isOutermost || _app == null)
        {
            return;
        }

        try
        {
            if (_screenUpdatingSuppressed)
            {
                _app.ScreenUpdating = true;
            }
        }
        catch (COMException ex)
        {
            _logger.LogWarning(ex,
                "ExcelWriteGuard: Failed to restore ScreenUpdating (HResult: 0x{HResult:X8})",
                ex.HResult);
        }
        catch (InvalidComObjectException ex)
        {
            _logger.LogWarning(ex, "ExcelWriteGuard: COM object already released during ScreenUpdating restoration");
        }
    }
}
