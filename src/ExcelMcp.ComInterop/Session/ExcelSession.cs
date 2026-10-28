using System.Diagnostics.CodeAnalysis;

namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Main entry point for Excel COM interop operations using batch pattern.
/// All operations execute on dedicated STA threads with proper COM cleanup.
/// </summary>
public static class ExcelSession
{
    internal static TimeSpan? TemporaryBatchTimeoutForTests { get; set; }
    internal static TimeSpan? TemporaryBatchStartupTimeoutForTests { get; set; }
    internal static Action<ExcelBatch>? TemporaryBatchCreatedForTests { get; set; }
    private static readonly AsyncLocal<Func<string[], bool, TimeSpan?, TimeSpan?, ExcelBatch>?> TemporaryBatchOwner = new();

    internal static bool HasTemporaryBatchOwner => TemporaryBatchOwner.Value is not null;

    internal static IDisposable OwnTemporaryBatches(
        Func<string[], bool, TimeSpan?, TimeSpan?, ExcelBatch> create)
    {
        var previous = TemporaryBatchOwner.Value;
        TemporaryBatchOwner.Value = create;
        return new TemporaryBatchOwnerScope(previous);
    }

    private sealed class TemporaryBatchOwnerScope(
        Func<string[], bool, TimeSpan?, TimeSpan?, ExcelBatch>? previous) : IDisposable
    {
        public void Dispose() => TemporaryBatchOwner.Value = previous;
    }

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
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Excel desktop automation is Windows-only; the shared BeginBatchWithTimeouts boundary enforces the platform requirement.")]
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
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "Excel desktop automation is Windows-only; the shared BeginBatchWithTimeouts boundary enforces the platform requirement.")]
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
        => CreateBatch(show, operationTimeout, startupTimeout, false, filePaths);

    internal static ExcelBatch BeginOwnedBatch(
        bool show, TimeSpan? operationTimeout, TimeSpan? startupTimeout, string filePath)
        => CreateBatch(show, operationTimeout, startupTimeout, true, filePath);

    private static ExcelBatch CreateBatch(
        bool show, TimeSpan? operationTimeout, TimeSpan? startupTimeout,
        bool deferStartup, params string[] filePaths)
    {
        if (filePaths == null || filePaths.Length == 0)
            throw new ArgumentException("At least one file path is required", nameof(filePaths));

        if (!deferStartup && TemporaryBatchOwner.Value is { } owner)
            return owner(filePaths.Select(path => WorkbookLocation.IsRemote(path)
                ? path : WorkbookLocation.Normalize(path)).ToArray(), show, operationTimeout, startupTimeout);

        // Create batch - it will create Excel/workbook on its own STA thread
        var batch = new ExcelBatch(
            filePaths,
            logger: null,
            show: show,
            operationTimeout: operationTimeout ?? (!deferStartup ? TemporaryBatchTimeoutForTests : null),
            startupTimeout: startupTimeout ?? (!deferStartup ? TemporaryBatchStartupTimeoutForTests : null) ?? (!deferStartup && TemporaryBatchTimeoutForTests is not null
                ? ComInteropConstants.DefaultOperationTimeout : null),
            deferStartup: true);
        if (!deferStartup)
        {
            TemporaryBatchCreatedForTests?.Invoke(batch);
            batch.WaitForStartup();
        }
        return batch;
    }

    /// <summary>
    /// Opens one workbook read-only for validation and returns a disposable batch.
    /// </summary>
    internal static ExcelBatch BeginReadOnlyValidation(
        string filePath,
        TimeSpan? operationTimeout,
        bool deferStartup = false)
    {
        return new ExcelBatch(
            [filePath],
            logger: null,
            show: WorkbookLocation.IsRemote(filePath),
            operationTimeout: operationTimeout,
            openReadOnly: true,
            deferStartup: deferStartup);
    }

}
