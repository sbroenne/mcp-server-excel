using System.Diagnostics.CodeAnalysis;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Main entry point for Excel COM interop operations using batch pattern.
/// All operations execute on dedicated STA threads with proper COM cleanup.
/// </summary>
public static class ExcelSession
{
    /// <summary>
    /// Global lock to serialize file creation operations.
    /// Prevents resource exhaustion from parallel CreateNew() calls.
    /// Each CreateNew() spawns a temporary Excel instance - must be sequential.
    /// </summary>
    private static readonly SemaphoreSlim _createFileLock = new(1, 1);
    /// <summary>
    /// Begins a batch of Excel operations against a single workbook instance.
    /// The Excel instance remains open until the batch is disposed, enabling multiple operations
    /// without incurring Excel startup/shutdown overhead.
    /// </summary>
    /// <param name="filePath">Path to the Excel file</param>
    /// <returns>IExcelBatch for executing multiple operations</returns>
    /// <remarks>
    /// All CLI and MCP operations use this batch-based approach for optimal performance.
    ///
    /// <para><b>Example:</b></para>
    /// <code>
    /// using var batch = ExcelSession.BeginBatch(filePath);
    ///
    /// // Synchronous COM operations
    /// batch.Execute((ctx, ct) => {
    ///     ctx.Book.Worksheets.Add("Sales");
    ///     return 0;
    /// });
    ///
    /// batch.Execute((ctx, ct) => {
    ///     ctx.Book.Worksheets.Add("Expenses");
    ///     return 0;
    /// });
    ///
    /// // Explicit save
    /// batch.Save();
    ///
    /// // Dispose closes workbook and quits Excel
    /// </code>
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static IExcelBatch BeginBatch(
        string filePath,
        CancellationToken _ = default)
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
        return new ExcelBatch(fullPath);
    }

    /// <summary>
    /// Creates a new Excel workbook at the specified path with a synchronous COM operation.
    /// Creates a minimal workbook then allows executing an operation before saving.
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="filePath">Path where to save the new Excel file</param>
    /// <param name="isMacroEnabled">Whether to create a macro-enabled workbook (.xlsm)</param>
    /// <param name="operation">Synchronous COM operation to execute with ExcelContext</param>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>Result of the operation</returns>
    /// <remarks>
    /// <para><b>File creation is automatically serialized</b> to prevent resource exhaustion.</para>
    /// <para>Even if called in parallel (e.g., Task.WhenAll), calls are queued and executed one at a time.</para>
    /// <para>This prevents spawning multiple temporary Excel.Application processes simultaneously.</para>
    /// </remarks>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static T CreateNew<T>(
        string filePath,
        bool isMacroEnabled,
        Func<ExcelContext, CancellationToken, T> operation,
        TimeSpan? _ = null,
        CancellationToken cancellationToken = default)
    {
        // CRITICAL: Acquire lock to serialize file creation operations
        // This prevents parallel CreateNew() calls from spawning multiple Excel processes
        _createFileLock.Wait(cancellationToken);
        try
        {
            string fullPath = Path.GetFullPath(filePath);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Create temporary empty workbook by launching Excel briefly
            Task.Run(() =>
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
                    // Close workbook and quit Excel
                    if (workbook != null)
                    {
                        try { workbook.Close(false); } catch { }
                        workbook = null; // Let GC handle RCW cleanup
                    }

                    if (excel != null)
                    {
                        // CodeQL suppression: Safe COM interop - excel is guaranteed to be Excel.Application type with Quit() method
                        // The dynamic type is intentional for late-binding COM automation
#pragma warning disable CS8602 // Dereference of a possibly null reference
                        try { excel.Quit(); } catch { }
#pragma warning restore CS8602
                        excel = null; // Let GC handle RCW cleanup
                    }
                    // The GC will automatically handle RCW cleanup through finalizers
                }
            }, cancellationToken).Wait(cancellationToken);

            // Now use batch API to execute the operation
            using var batch = BeginBatch(fullPath, cancellationToken);
            var result = batch.Execute(operation, cancellationToken);
            // Note: Caller is responsible for saving if needed

            return result;
        }
        finally
        {
            // Release lock to allow next CreateNew() call
            _createFileLock.Release();
        }
    }


}


