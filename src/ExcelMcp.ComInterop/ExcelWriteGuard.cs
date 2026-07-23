using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop;

/// <summary>
/// Provides automatic COM safety for Excel operations by suppressing events, screen updating,
/// and automatic calculation during the guard's lifetime. Restores original state on disposal.
///
/// This guard is integrated into <see cref="Session.ExcelBatch.Execute{T}"/> so ALL operations
/// get protection automatically — no manual suppression needed in command implementations.
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

    // Captured state (only set for outermost guard)
    private bool _originalScreenUpdating;

    private bool _disposed;

    /// <summary>
    /// Creates a new write guard that suppresses Excel events, screen updating, and calculation.
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
            // Capture current state
            _originalScreenUpdating = _app.ScreenUpdating;

            // Suppress screen updating universally — safe for ALL operations.
            // Prevents Excel from repainting during every COM call.
            //
            // NOTE: EnableEvents and Calculation are NOT suppressed here.
            // Suppressing events breaks Data Model operations (AddToDataModel, CreateRelationship,
            // CreateMeasure) which rely on internal Excel events for model synchronization.
            // Suppressing calculation breaks PivotTable refresh and Power Query refresh.
            // Commands that need these suppressions handle them individually.
            if (_originalScreenUpdating)
            {
                _app.ScreenUpdating = false;
            }
        }
        catch (COMException ex)
        {
            // Excel COM proxy may be disconnected (process died). Log and continue —
            // the operation will fail with its own exception; we just can't guard it.
            _logger.LogWarning(ex,
                "ExcelWriteGuard: Failed to capture/suppress Excel state (HResult: 0x{HResult:X8}). " +
                "Excel may have crashed.", ex.HResult);
        }
        catch (InvalidComObjectException ex)
        {
            _logger.LogWarning(ex, "ExcelWriteGuard: COM object already released");
        }
    }

    /// <summary>
    /// Restores original Excel state (events, screen updating, calculation).
    /// Only the outermost guard restores — inner guards are no-ops.
    /// </summary>
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
            if (_originalScreenUpdating)
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
        catch (InvalidComObjectException)
        {
            // COM proxy dead — nothing more to do
        }
    }
}
