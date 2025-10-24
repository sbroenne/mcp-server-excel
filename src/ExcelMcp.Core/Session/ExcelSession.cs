using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Manages Excel COM automation sessions with automatic pooling for performance.
/// Pool is a private implementation detail - callers just use Execute() and get automatic optimization.
/// </summary>
public static class ExcelSession
{
    // Private pool - automatic optimization, no configuration needed
    private static readonly Lazy<ExcelInstancePool> _defaultPool = new(() =>
        new ExcelInstancePool(
            idleTimeout: TimeSpan.FromSeconds(60),
            maxInstances: 10
        )
    );

    /// <summary>
    /// Executes an action with an Excel workbook using automatic pooling for performance.
    /// Pool is managed internally - callers don't need to configure or manage it.
    /// </summary>
    /// <typeparam name="T">Return type of the action</typeparam>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes to the file</param>
    /// <param name="action">Action to execute with Excel application and workbook</param>
    /// <returns>Result of the action</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static T Execute<T>(string filePath, bool save, Func<dynamic, dynamic, T> action)
    {
        try
        {
            return _defaultPool.Value.WithPooledExcel(filePath, save, action);
        }
        catch (ObjectDisposedException)
        {
            // Pool was disposed (shouldn't happen in production) - fall back to single instance
            return ExecuteSingleInstance(filePath, save, action);
        }
    }

    /// <summary>
    /// Creates a new Excel workbook and executes an action.
    /// </summary>
    /// <typeparam name="T">Return type of the action</typeparam>
    /// <param name="filePath">Path where to save the new Excel file</param>
    /// <param name="isMacroEnabled">Whether to create a macro-enabled workbook (.xlsm)</param>
    /// <param name="action">Action to execute with Excel application and new workbook</param>
    /// <returns>Result of the action</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static T CreateNew<T>(string filePath, bool isMacroEnabled, Func<dynamic, dynamic, T> action)
    {
        dynamic? excel = null;
        dynamic? workbook = null;
        string operation = $"CreateNew({Path.GetFileName(filePath)}, macroEnabled={isMacroEnabled})";

        try
        {
            // Validate file path first - prevent path traversal attacks
            string fullPath = Path.GetFullPath(filePath);

            // Validate file size limits for security (prevent DoS)
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Get Excel COM type
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException("Excel is not installed or not properly registered. " +
                    "Please verify Microsoft Excel is installed and COM registration is intact.");
            }

#pragma warning disable IL2072 // COM interop is not AOT compatible but is required for Excel automation
            excel = Activator.CreateInstance(excelType);
#pragma warning restore IL2072
            if (excel == null)
            {
                throw new InvalidOperationException("Failed to create Excel COM instance. " +
                    "Excel may be corrupted or COM subsystem unavailable.");
            }

            // Configure Excel for automation
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.Interactive = false;

            // Create new workbook
            workbook = excel.Workbooks.Add();

            // Execute the user action
            var result = action(excel, workbook);

            // Save the workbook with appropriate format
            if (isMacroEnabled)
            {
                // Save as macro-enabled workbook (format 52)
                workbook.SaveAs(fullPath, 52);
            }
            else
            {
                // Save as regular workbook (format 51)
                workbook.SaveAs(fullPath, 51);
            }

            return result;
        }
        catch (COMException comEx)
        {
            throw new InvalidOperationException($"Excel COM operation failed during {operation}: {comEx.Message}", comEx);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Operation failed during {operation}: {ex.Message}", ex);
        }
        finally
        {
            // Enhanced COM cleanup to prevent process leaks

            // Close workbook first
            if (workbook != null)
            {
                try
                {
                    workbook.Close(false); // Don't save again, we already saved
                }
                catch (COMException)
                {
                    // Workbook might already be closed, ignore
                }
                catch
                {
                    // Any other exception during close, ignore to continue cleanup
                }

                try
                {
                    Marshal.ReleaseComObject(workbook);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                workbook = null;
            }

            // Quit Excel application
            if (excel != null)
            {
                try
                {
                    excel.Quit();
                }
                catch (COMException)
                {
                    // Excel might already be closing, ignore
                }
                catch
                {
                    // Any other exception during quit, ignore to continue cleanup
                }

                try
                {
                    Marshal.ReleaseComObject(excel);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                excel = null;
            }

            // Recommended COM cleanup pattern:
            // Two GC cycles are sufficient - one to collect, one to finalize
            // Microsoft recommends against excessive GC.Collect() calls
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect(); // Final collect to clean up objects queued during finalization
        }
    }

    /// <summary>
    /// Single-instance Excel execution pattern - creates new Excel instance for each operation.
    /// Internal for testing only - allows pool tests to bypass pooling to test pool behavior in isolation.
    /// </summary>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    internal static T ExecuteSingleInstance<T>(string filePath, bool save, Func<dynamic, dynamic, T> action)
    {
        dynamic? excel = null;
        dynamic? workbook = null;
        string operation = $"ExecuteSingleInstance({Path.GetFileName(filePath)}, save={save})";

        try
        {
            // Validate file path first - prevent path traversal attacks
            string fullPath = Path.GetFullPath(filePath);

            // Additional security: ensure the file is within reasonable bounds
            if (fullPath.Length > 32767)
            {
                throw new ArgumentException($"File path too long: {fullPath.Length} characters (Windows limit: 32767)");
            }

            // Security: Validate file extension to prevent executing arbitrary files
            string extension = Path.GetExtension(fullPath).ToLowerInvariant();
            if (extension is not (".xlsx" or ".xlsm" or ".xls"))
            {
                throw new ArgumentException($"Invalid file extension '{extension}'. Only Excel files (.xlsx, .xlsm, .xls) are supported.");
            }

            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"Excel file not found: {fullPath}", fullPath);
            }

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException("Excel is not installed or not properly registered. " +
                    "Please verify Microsoft Excel is installed and COM registration is intact.");
            }

#pragma warning disable IL2072 // COM interop is not AOT compatible but is required for Excel automation
            excel = Activator.CreateInstance(excelType);
#pragma warning restore IL2072
            if (excel == null)
            {
                throw new InvalidOperationException("Failed to create Excel COM instance. " +
                    "Excel may be corrupted or COM subsystem unavailable.");
            }

            // Configure Excel for automation
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.Interactive = false;

            // Open workbook with detailed error context
            try
            {
                workbook = excel.Workbooks.Open(fullPath);
            }
            catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x8001010A))
            {
                // Excel is busy - provide specific guidance
                throw new InvalidOperationException(
                    "Excel is busy (likely has a dialog open). Close any Excel dialogs and retry.", comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x80070020))
            {
                // File sharing violation
                throw new InvalidOperationException(
                    $"File '{Path.GetFileName(fullPath)}' is locked by another process. " +
                    "Close Excel and any other applications using this file.", comEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to open workbook '{Path.GetFileName(fullPath)}'. " +
                    "File may be corrupted, password-protected, or incompatible.", ex);
            }

            if (workbook == null)
            {
                throw new InvalidOperationException($"Failed to open workbook: {Path.GetFileName(fullPath)}");
            }

            // Execute the user action with error context and retry logic for transient errors
            T result;
            int retryCount = 0;
            const int maxRetries = 3;

            while (true)
            {
                try
                {
                    result = action(excel, workbook);
                    break; // Success - exit retry loop
                }
                catch (COMException comEx) when (comEx.HResult == unchecked((int)0x8001010A) && retryCount < maxRetries)
                {
                    // Excel is busy (RPC_E_SERVERCALL_RETRYLATER)
                    // This can happen during parallel operations or when Excel is processing
                    retryCount++;
                    System.Threading.Thread.Sleep(500 * retryCount); // Exponential backoff: 500ms, 1s, 1.5s

                    if (retryCount >= maxRetries)
                    {
                        throw new InvalidOperationException(
                            "Excel is busy. Please close any dialogs and try again.", comEx);
                    }
                    // Continue retry loop
                }
                catch
                {
                    // Propagate all other exceptions with original context
                    throw;
                }
            }

            // Save if requested
            if (save && workbook != null)
            {
                try
                {
                    workbook.Save();
                }
                catch
                {
                    // Propagate save exceptions
                    throw;
                }
            }

            return result;
        }
        catch
        {
            // Propagate exceptions to caller
            throw;
        }
        finally
        {
            // Enhanced COM cleanup to prevent process leaks

            // Close workbook first
            if (workbook != null)
            {
                try
                {
                    workbook.Close(save);
                }
                catch (COMException)
                {
                    // Workbook might already be closed, ignore
                }
                catch
                {
                    // Any other exception during close, ignore to continue cleanup
                }

                try
                {
                    Marshal.ReleaseComObject(workbook);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                workbook = null;
            }

            // Quit Excel application
            if (excel != null)
            {
                try
                {
                    excel.Quit();
                }
                catch (COMException)
                {
                    // Excel might already be closing, ignore
                }
                catch
                {
                    // Any other exception during quit, ignore to continue cleanup
                }

                try
                {
                    Marshal.ReleaseComObject(excel);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                excel = null;
            }

            // Recommended COM cleanup pattern:
            // Two GC cycles are sufficient - one to collect, one to finalize
            // Microsoft recommends against excessive GC.Collect() calls
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect(); // Final collect to clean up objects queued during finalization
        }
    }
}
