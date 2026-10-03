using System.Diagnostics.CodeAnalysis;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Main entry point for Excel COM interop operations using batch pattern.
/// All operations execute on dedicated STA threads with proper COM cleanup.
/// </summary>
public static class ExcelSession
{
    /// <summary>
    /// Begins a batch of Excel operations against one or more workbook instances.
    /// The Excel instance remains open until the batch is disposed, enabling multiple operations
    /// without incurring Excel startup/shutdown overhead.
    /// </summary>
    /// <param name="filePaths">Paths to Excel files. First file is the primary workbook.</param>
    /// <returns>IExcelBatch for executing multiple operations</returns>
    /// <remarks>
    /// All CLI and MCP operations use this batch-based approach for optimal performance.
    /// For cross-workbook operations (copy, move), pass multiple file paths.
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
    public static IExcelBatch BeginBatch(params string[] filePaths)
        => BeginBatch(show: false, operationTimeout: null, filePaths);

    /// <summary>
    /// Begins a batch of Excel operations against one or more workbook instances with optional UI visibility.
    /// The Excel instance remains open until the batch is disposed, enabling multiple operations
    /// without incurring Excel startup/shutdown overhead.
    /// </summary>
    /// <param name="show">Whether to show the Excel window (default: false for background automation).</param>
    /// <param name="operationTimeout">Maximum time for startup and any single operation (default: 120 seconds).</param>
    /// <param name="filePaths">Paths to Excel files. First file is the primary workbook.</param>
    /// <returns>IExcelBatch for executing multiple operations</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static IExcelBatch BeginBatch(
        bool show,
        TimeSpan? operationTimeout,
        params string[] filePaths)
        => BeginBatchWithTimeouts(show, operationTimeout, operationTimeout, filePaths);

    /// <summary>
    /// Test-only seam that preserves the production timeout contract while allowing
    /// operation timeout regressions to use the normal Excel startup allowance.
    /// </summary>
    internal static IExcelBatch BeginBatchWithTimeouts(
        bool show,
        TimeSpan? operationTimeout,
        TimeSpan? startupTimeout,
        params string[] filePaths)
    {
        if (filePaths == null || filePaths.Length == 0)
            throw new ArgumentException("At least one file path is required", nameof(filePaths));

        string[] fullPaths = new string[filePaths.Length];
        for (int i = 0; i < filePaths.Length; i++)
        {
            string fullPath = Path.GetFullPath(filePaths[i]);

            // Validate file exists
            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"Excel file not found: {fullPath}. To create a new file, use the 'create' action instead of 'open'.", fullPath);
            }

            // Security: Validate file extension
            string extension = Path.GetExtension(fullPath).ToLowerInvariant();
            if (extension is not (".xlsx" or ".xlsm" or ".xlsb" or ".xls"))
            {
                throw new ArgumentException($"Invalid file extension '{extension}'. Only Excel files (.xlsx, .xlsm, .xlsb, .xls) are supported.");
            }

            fullPaths[i] = fullPath;
        }

        // Create batch - it will create Excel/workbook on its own STA thread
        return new ExcelBatch(
            fullPaths,
            logger: null,
            show: show,
            operationTimeout: operationTimeout,
            startupTimeout: startupTimeout);
    }

    /// <summary>
    /// Opens one workbook read-only for validation and returns a disposable batch.
    /// </summary>
    internal static IExcelBatch BeginReadOnlyValidation(
        string filePath,
        TimeSpan? operationTimeout)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException(
                $"Excel file not found: {fullPath}.",
                fullPath);
        }

        var extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not (".xlsx" or ".xlsm" or ".xlsb" or ".xls"))
        {
            throw new ArgumentException(
                $"Invalid file extension '{extension}'. Validation supports .xlsx, .xlsm, .xlsb and .xls only.",
                nameof(filePath));
        }

        return new ExcelBatch(
            [fullPath],
            logger: null,
            show: false,
            operationTimeout: operationTimeout,
            openReadOnly: true);
    }

}
