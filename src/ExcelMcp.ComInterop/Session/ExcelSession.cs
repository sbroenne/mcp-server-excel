using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Main entry point for Excel COM interop operations using batch pattern.
/// All operations execute on dedicated STA threads with proper COM cleanup.
/// </summary>
public static class ExcelSession
{
    /// <summary>
    /// Begins a batch of Excel operations against a single workbook instance.
    /// The Excel instance remains open until the batch is disposed, enabling multiple operations
    /// without incurring Excel startup/shutdown overhead.
    /// </summary>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="cancellationToken">Optional cancellation token</param>
    /// <returns>IExcelBatch for executing multiple operations</returns>
    /// <remarks>
    /// All CLI and MCP operations use this batch-based approach for optimal performance.
    ///
    /// <para><b>Example:</b></para>
    /// <code>
    /// await using var batch = await ExcelSession.BeginBatchAsync(filePath);
    ///
    /// // Synchronous COM operations
    /// await batch.Execute((ctx, ct) => {
    ///     ctx.Book.Worksheets.Add("Sales");
    ///     return 0;
    /// });
    ///
    /// await batch.Execute((ctx, ct) => {
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
    /// Creates a new Excel workbook at the specified path with a synchronous COM operation.
    /// Creates a minimal workbook then allows executing an operation before saving.
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="filePath">Path where to save the new Excel file</param>
    /// <param name="isMacroEnabled">Whether to create a macro-enabled workbook (.xlsm)</param>
    /// <param name="operation">Synchronous COM operation to execute with ExcelContext</param>
    /// <param name="timeout">Optional timeout for the operation (ignored for now, reserved for future use)</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Result of the operation</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static async Task<T> CreateNew<T>(
        string filePath,
        bool isMacroEnabled,
        Func<ExcelContext, CancellationToken, T> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        string fullPath = Path.GetFullPath(filePath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Create temporary empty workbook by launching Excel briefly
        await Task.Run(() =>
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

                excel.Visible = false;
                excel.DisplayAlerts = false;

                workbook = excel.Workbooks.Add();

                // Save immediately to create the file
                if (isMacroEnabled)
                {
                    workbook.SaveAs(fullPath, 52); // xlOpenXMLWorkbookMacroEnabled
                }
                else
                {
                    workbook.SaveAs(fullPath, 51); // xlOpenXMLWorkbook
                }
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
                    // CodeQL suppression: Safe COM interop - excel is guaranteed to be Excel.Application type with Quit() method
                    // The dynamic type is intentional for late-binding COM automation
#pragma warning disable CS8602 // Dereference of a possibly null reference
                    try { excel.Quit(); } catch { }
#pragma warning restore CS8602
                    try { Marshal.FinalReleaseComObject(excel); } catch { }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        });

        // Now use batch API to execute the operation
        await using var batch = await BeginBatchAsync(fullPath, cancellationToken);
        var result = await batch.Execute(operation);
        await batch.SaveAsync();

        return result;
    }

    /// <summary>
    /// Creates a new Excel workbook at the specified path with an async operation.
    /// Creates a minimal workbook then allows executing an operation before saving.
    /// Use this ONLY when the operation performs async I/O (file operations).
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
        Func<ExcelContext, CancellationToken, Task<T>> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default)
    {
        string fullPath = Path.GetFullPath(filePath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Create temporary empty workbook by launching Excel briefly
        await Task.Run(() =>
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

                excel.Visible = false;
                excel.DisplayAlerts = false;

                workbook = excel.Workbooks.Add();

                // Save immediately to create the file
                if (isMacroEnabled)
                {
                    workbook.SaveAs(fullPath, 52); // xlOpenXMLWorkbookMacroEnabled
                }
                else
                {
                    workbook.SaveAs(fullPath, 51); // xlOpenXMLWorkbook
                }
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
                    // CodeQL suppression: Safe COM interop - excel is guaranteed to be Excel.Application type with Quit() method
                    // The dynamic type is intentional for late-binding COM automation
#pragma warning disable CS8602 // Dereference of a possibly null reference
                    try { excel.Quit(); } catch { }
#pragma warning restore CS8602
                    try { Marshal.FinalReleaseComObject(excel); } catch { }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        });

        // Now use batch API to execute the operation
        await using var batch = await BeginBatchAsync(fullPath, cancellationToken);
        var result = await batch.ExecuteAsync(operation);
        await batch.SaveAsync();

        return result;
    }
}
