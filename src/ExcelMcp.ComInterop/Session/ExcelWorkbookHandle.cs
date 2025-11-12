using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Wraps Excel COM objects for a single workbook.
/// Tracks last access time for automatic cleanup.
/// Handles COM object lifecycle and disposal.
/// </summary>
public sealed class ExcelWorkbookHandle : IAsyncDisposable, IDisposable
{
    private dynamic? _application;
    private dynamic? _workbook;
    private bool _disposed;

    /// <summary>
    /// Gets the absolute file path to the workbook.
    /// </summary>
    public string FilePath { get; }

    /// <summary>
    /// Gets the last access time (UTC) for this handle.
    /// Updated on each operation to track inactivity.
    /// </summary>
    public DateTime LastAccess { get; private set; }

    /// <summary>
    /// Gets the Excel.Application COM object.
    /// </summary>
    /// <exception cref="ObjectDisposedException">If handle has been disposed</exception>
    public dynamic Application => _application ?? throw new ObjectDisposedException(nameof(ExcelWorkbookHandle));

    /// <summary>
    /// Gets the Excel.Workbook COM object.
    /// </summary>
    /// <exception cref="ObjectDisposedException">If handle has been disposed</exception>
    public dynamic Workbook => _workbook ?? throw new ObjectDisposedException(nameof(ExcelWorkbookHandle));

    private ExcelWorkbookHandle(string filePath)
    {
        FilePath = Path.GetFullPath(filePath);
        LastAccess = DateTime.UtcNow;
    }

    /// <summary>
    /// Creates a new handle by opening an existing workbook file.
    /// </summary>
    /// <param name="filePath">Absolute or relative path to the workbook file</param>
    /// <returns>A new handle with Excel opened and workbook loaded</returns>
    /// <exception cref="InvalidOperationException">If Excel is not installed</exception>
    /// <exception cref="FileNotFoundException">If workbook file doesn't exist</exception>
    public static async Task<ExcelWorkbookHandle> OpenAsync(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);

        if (!File.Exists(absolutePath))
        {
            throw new FileNotFoundException(
                $"Workbook file not found: {absolutePath}",
                absolutePath);
        }

        var handle = new ExcelWorkbookHandle(absolutePath);

        await Task.Run(() =>
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException(
                    "Microsoft Excel is not installed on this system.");
            }

            handle._application = Activator.CreateInstance(excelType);
            handle._application.Visible = false;
            handle._application.DisplayAlerts = false;

            // Disable macro security warnings for unattended automation
            // msoAutomationSecurityForceDisable = 3 (disable all macros, no prompts)
            handle._application.AutomationSecurity = 3;

            // CRITICAL: Check if file is locked at OS level BEFORE attempting Excel COM open
            FileAccessValidator.ValidateFileNotLocked(absolutePath);

            try
            {
                handle._workbook = handle._application.Workbooks.Open(absolutePath);
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
            {
                // Excel Error 1004 - File is already open or locked
                throw FileAccessValidator.CreateFileLockedError(absolutePath, ex);
            }
        });

        return handle;
    }

    /// <summary>
    /// Creates a new workbook file and returns a handle to it.
    /// </summary>
    /// <param name="filePath">Path where the new workbook should be created</param>
    /// <returns>A new handle with Excel opened and new workbook created</returns>
    /// <exception cref="InvalidOperationException">If Excel is not installed or file already exists</exception>
    public static async Task<ExcelWorkbookHandle> CreateAsync(string filePath)
    {
        string absolutePath = Path.GetFullPath(filePath);

        if (File.Exists(absolutePath))
        {
            throw new InvalidOperationException(
                $"File already exists: {absolutePath}");
        }

        var handle = new ExcelWorkbookHandle(absolutePath);

        await Task.Run(() =>
        {
            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException(
                    "Microsoft Excel is not installed on this system.");
            }

            handle._application = Activator.CreateInstance(excelType);
            handle._application.Visible = false;
            handle._application.DisplayAlerts = false;
            handle._application.AutomationSecurity = 3;

            handle._workbook = handle._application.Workbooks.Add();
        });

        return handle;
    }

    /// <summary>
    /// Updates the last access time to current UTC time.
    /// Call this on each operation to reset inactivity timeout.
    /// </summary>
    public void UpdateLastAccess()
    {
        LastAccess = DateTime.UtcNow;
    }

    /// <summary>
    /// Saves the workbook to disk.
    /// </summary>
    public async Task SaveAsync()
    {
        ObjectDisposedException.ThrowIf(_disposed, nameof(ExcelWorkbookHandle));

        await Task.Run(() =>
        {
            // If this is a new workbook (never saved), use SaveAs
            // Otherwise use Save
            string workbookPath = _workbook.FullName?.ToString() ?? string.Empty;

            if (string.IsNullOrEmpty(workbookPath) ||
                workbookPath == "Book1" ||
                !File.Exists(workbookPath))
            {
                // New workbook - save with filePath
                _workbook.SaveAs(FilePath);
            }
            else
            {
                // Existing workbook - just save
                _workbook.Save();
            }
        });

        UpdateLastAccess();
    }

    /// <summary>
    /// Asynchronously disposes the handle, closing the workbook and Excel application.
    /// </summary>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        await Task.Run(() =>
        {
            try
            {
                // Close workbook without saving
                if (_workbook != null)
                {
                    try
                    {
                        _workbook.Close(false);
                    }
                    catch
                    {
                        // Suppress errors during close
                    }
                    ComUtilities.Release(ref _workbook);
                }

                // Quit Excel
                if (_application != null)
                {
                    try
                    {
                        _application.Quit();
                    }
                    catch
                    {
                        // Suppress errors during quit
                    }
                    ComUtilities.Release(ref _application);
                }

                // Force COM cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch
            {
                // Suppress all disposal errors to prevent exceptions during dispose
            }
        });
    }

    /// <summary>
    /// Synchronously disposes the handle.
    /// </summary>
    public void Dispose()
    {
        DisposeAsync().AsTask().GetAwaiter().GetResult();
    }
}
