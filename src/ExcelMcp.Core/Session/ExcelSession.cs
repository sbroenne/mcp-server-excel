using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Core.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Session;

/// <summary>
/// Main entry point for Excel COM interop operations using async/await pattern.
/// Provides single-instance execution for atomic operations and batch mode for workflows.
/// </summary>
public static class ExcelSession
{
    static ExcelSession()
    {
        // OleMessageFilter registration removed - not essential for this version
        // Can be re-added later if needed for better COM error handling
    }

    /// <summary>
    /// Executes an async operation on an Excel workbook.
    /// Opens Excel, executes one operation, optionally saves, then closes immediately.
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes to the file</param>
    /// <param name="operation">Async operation to execute with ExcelContext</param>
    /// <param name="timeout">Optional timeout for the operation (ignored for now, reserved for future use)</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Result of the operation</returns>
    /// <remarks>
    /// This is the standard method for CLI atomic operations and single MCP operations.
    /// For multiple operations on the same workbook, use BeginBatchAsync() instead.
    /// </remarks>
    public static async Task<T> ExecuteAsync<T>(
        string filePath,
        bool save,
        Func<ExcelContext, CancellationToken, ValueTask<T>> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        // Validate file exists
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Excel file not found: {filePath}", filePath);
        }

        return await ExecuteSingleInstanceAsync(filePath, save, operation, timeout, cancellationToken);
    }

    /// <summary>
    /// Begins a batch of Excel operations against a single workbook instance.
    /// The Excel instance remains open until the batch is disposed, enabling multiple operations
    /// without incurring Excel startup/shutdown overhead.
    /// </summary>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="cancellationToken">Optional cancellation token</param>
    /// <returns>IExcelBatch for executing multiple operations</returns>
    /// <remarks>
    /// Use this for MCP Server multi-operation workflows where you need to execute multiple
    /// commands against the same workbook efficiently. For single operations (CLI, standalone
    /// MCP operations), use ExecuteAsync() instead.
    ///
    /// <para><b>Example:</b></para>
    /// <code>
    /// await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    ///
    /// // Operation 1
    /// await batch.ExecuteAsync(async (ctx, ct) => {
    ///     ctx.Book.Worksheets.Add("Sales");
    ///     return 0;
    /// });
    ///
    /// // Operation 2
    /// await batch.ExecuteAsync(async (ctx, ct) => {
    ///     ctx.Book.Worksheets.Add("Expenses");
    ///     return 0;
    /// });
    ///
    /// // Explicit save
    /// await batch.SaveAsync();
    ///
    /// // Dispose closes workbook and quits Excel
    /// </code>
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static async Task<IExcelBatch> BeginBatchAsync(
        string filePath,
        CancellationToken cancellationToken = default)
    {
        string fullPath = Path.GetFullPath(filePath);

        // Validate file exists
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Excel file not found: {fullPath}", fullPath);
        }

        // Security: Validate file extension
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not (".xlsx" or ".xlsm" or ".xls"))
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only Excel files (.xlsx, .xlsm, .xls) are supported.");
        }

        // Create batch - it will create Excel/workbook on its own STA thread
        return await Task.FromResult<IExcelBatch>(new ExcelBatch(fullPath));
    }

    /// <summary>
    /// Executes operation using a single Excel instance (atomic operation pattern).
    /// Opens Excel, executes one operation, optionally saves, then closes and quits.
    /// CRITICAL: All COM operations execute on a dedicated STA thread with OLE message filter.
    /// </summary>
    private static async Task<T> ExecuteSingleInstanceAsync<T>(
        string filePath,
        bool save,
        Func<ExcelContext, CancellationToken, ValueTask<T>> operation,
        TimeSpan? timeout,
        CancellationToken cancellationToken)
    {
        string fullPath = Path.GetFullPath(filePath);

        // Validate file exists
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Excel file not found: {fullPath}", fullPath);
        }

        // Security: Validate file extension
        string extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not (".xlsx" or ".xlsm" or ".xls"))
        {
            throw new ArgumentException($"Invalid file extension '{extension}'. Only Excel files (.xlsx, .xlsm, .xls) are supported.");
        }

        // CRITICAL: Execute all COM operations on STA thread with OLE message filter
        return await ExcelStaExecutor.ExecuteOnStaThreadAsync(async () =>
        {
            dynamic? excel = null;
            dynamic? workbook = null;

            try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException("Excel is not installed or not properly registered.");
            }

#pragma warning disable IL2072
            excel = Activator.CreateInstance(excelType);
#pragma warning restore IL2072

            if (excel == null)
            {
                throw new InvalidOperationException("Failed to create Excel COM instance.");
            }

            // Configure Excel for automation
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.Interactive = false;

            // Open workbook
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

            if (workbook == null)
            {
                throw new InvalidOperationException($"Failed to open workbook: {Path.GetFileName(fullPath)}");
            }

            // Execute operation with context
            var context = new ExcelContext(fullPath, excel, workbook);
            var result = await operation(context, cancellationToken);

            // Save if requested
            if (save)
            {
                try
                {
                    workbook.Save();
                }
                catch (COMException ex)
                {
                    // Common error codes:
                    // 0x800A03EC = VBA Error 1004 (general save error, often "read-only")
                    // 0x800AC472 = The file is locked for editing

                    if (ex.Message.Contains("read-only") ||
                        ex.HResult == unchecked((int)0x800A03EC) ||
                        ex.HResult == unchecked((int)0x800AC472))
                    {
                        // Try SaveAs as a workaround
                        int fileFormat = extension == ".xlsm" ? 52 : 51;
                        workbook.SaveAs(fullPath, fileFormat);
                    }
                    else
                    {
                        throw;
                    }
                }
            }

            return result;
        }
        finally
        {
            // Enhanced COM cleanup to prevent process leaks

            // Close workbook first
            if (workbook != null)
            {
                try
                {
                    workbook.Close(false); // Don't save on close - save was explicit above
                }
                catch
                {
                    // Workbook might already be closed, ignore to continue cleanup
                }

                try
                {
                    Marshal.FinalReleaseComObject(workbook);
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
                catch
                {
                    // Excel might already be closing, ignore to continue cleanup
                }

                try
                {
                    Marshal.FinalReleaseComObject(excel);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                excel = null;
            }

            // Recommended COM cleanup pattern:
            // Two GC cycles are sufficient - one to collect, one to finalize
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect(); // Final collect to clean up objects queued during finalization
            }
        }, cancellationToken); // Close ExcelStaExecutor.ExecuteOnStaThreadAsync
    }

    /// <summary>
    /// Creates a new Excel workbook and executes an async operation.
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="filePath">Path where to save the new Excel file</param>
    /// <param name="isMacroEnabled">Whether to create a macro-enabled workbook (.xlsm)</param>
    /// <param name="operation">Async operation to execute with ExcelContext</param>
    /// <param name="timeout">Optional timeout for the operation (ignored for now, reserved for future use)</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Result of the operation</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static async Task<T> CreateNewAsync<T>(
        string filePath,
        bool isMacroEnabled,
        Func<ExcelContext, CancellationToken, ValueTask<T>> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        string fullPath = Path.GetFullPath(filePath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // CRITICAL: Execute all COM operations on STA thread with OLE message filter
        return await ExcelStaExecutor.ExecuteOnStaThreadAsync(async () =>
        {
            dynamic? excel = null;
            dynamic? workbook = null;

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel is not installed or not properly registered.");
                }

#pragma warning disable IL2072
                excel = Activator.CreateInstance(excelType);
#pragma warning restore IL2072

                if (excel == null)
                {
                    throw new InvalidOperationException("Failed to create Excel COM instance.");
                }

                excel.Visible = false;
                excel.DisplayAlerts = false;
                excel.ScreenUpdating = false;
                excel.Interactive = false;

                workbook = excel.Workbooks.Add();

                var context = new ExcelContext(fullPath, excel, workbook);
                var result = await operation(context, cancellationToken);

                if (isMacroEnabled)
                {
                    workbook.SaveAs(fullPath, 52);
                }
                else
                {
                    workbook.SaveAs(fullPath, 51);
                }

                return result;
            }
            finally
            {
                if (workbook != null)
                {
                    try { workbook.Close(false); } catch { }
                    try { Marshal.FinalReleaseComObject(workbook); } catch { }
                }

                if (excel != null)
                {
                    try { excel.Quit(); } catch { }
                    try { Marshal.FinalReleaseComObject(excel); } catch { }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }, cancellationToken); // Close ExcelStaExecutor.ExecuteOnStaThreadAsync
    }
}
