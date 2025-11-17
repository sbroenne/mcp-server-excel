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

            CreateWorkbookOnStaThread(fullPath, isMacroEnabled, cancellationToken);

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

    private static void CreateWorkbookOnStaThread(string fullPath, bool isMacroEnabled, CancellationToken cancellationToken)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            dynamic? excel = null;
            dynamic? workbook = null;

            try
            {
                OleMessageFilter.Register();

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

                if (isMacroEnabled)
                {
                    workbook.SaveAs(fullPath, 52); // xlOpenXMLWorkbookMacroEnabled
                }
                else
                {
                    workbook.SaveAs(fullPath, 51); // xlOpenXMLWorkbook
                }

                completion.SetResult();
            }
            catch (Exception ex)
            {
                completion.TrySetException(ex);
            }
            finally
            {
                if (workbook != null)
                {
                    try { workbook.Close(false); }
                    catch { }

                    workbook = null;
                }

                if (excel != null)
                {
                    try { excel.Quit(); }
                    catch { }

                    excel = null;
                }

                OleMessageFilter.Revoke();
            }
        })
        {
            IsBackground = true,
            Name = $"ExcelCreate-{Path.GetFileName(fullPath)}"
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        completion.Task.Wait(cancellationToken);

        // Ensure thread finished before proceeding
        thread.Join();
    }
}


