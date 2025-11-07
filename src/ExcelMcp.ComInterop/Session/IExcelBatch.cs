namespace Sbroenne.ExcelMcp.ComInterop.Session;

/// <summary>
/// Represents a batch of Excel operations that share a single Excel instance.
/// Implements IAsyncDisposable to ensure proper COM cleanup.
/// </summary>
/// <remarks>
/// Use this interface via ExcelSession.BeginBatchAsync() for multi-operation workflows.
/// The batch keeps Excel and the workbook open until disposed, enabling efficient
/// execution of multiple commands without repeated Excel startup/shutdown overhead.
///
/// <para><b>Lifecycle:</b></para>
/// <list type="bullet">
/// <item>Created via ExcelSession.BeginBatchAsync(filePath)</item>
/// <item>Operations executed via ExecuteAsync()</item>
/// <item>Optional explicit save via SaveAsync()</item>
/// <item>Disposed via DisposeAsync() or "await using" pattern</item>
/// </list>
///
/// <para><b>Example:</b></para>
/// <code>
/// await using var batch = await ExcelSession.BeginBatchAsync("workbook.xlsx");
///
/// // Synchronous COM operations
/// await batch.Execute((ctx, ct) => {
///     ctx.Book.Worksheets.Add("Data");
///     return 0;
/// });
///
/// await batch.Execute((ctx, ct) => {
///     ctx.Book.Worksheets["Data"].Range["A1"].Value = "Header";
///     return 0;
/// });
///
/// // Async I/O operations
/// await batch.ExecuteAsync(async (ctx, ct) => {
///     string formula = ctx.Book.Range["A1"].Formula;
///     await File.WriteAllTextAsync("output.txt", formula, ct);
///     return 0;
/// });
///
/// // Explicit save
/// await batch.SaveAsync();
/// </code>
/// </remarks>
public interface IExcelBatch : IAsyncDisposable
{
    /// <summary>
    /// Gets the path to the Excel workbook this batch operates on.
    /// </summary>
    string WorkbookPath { get; }

    /// <summary>
    /// Executes a synchronous COM operation within this batch.
    /// The operation receives an ExcelContext with access to the Excel app and workbook.
    /// Use this for pure COM operations (property access, method calls).
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="operation">Synchronous COM operation to execute</param>
    /// <param name="timeout">Optional timeout override. If not specified, uses default (2 minutes). Maximum is 5 minutes.</param>
    /// <param name="cancellationToken">Optional cancellation token</param>
    /// <returns>Result of the operation</returns>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Excel COM error occurred</exception>
    /// <exception cref="TimeoutException">Operation exceeded the timeout period</exception>
    Task<T> Execute<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes a genuinely async operation within this batch.
    /// The operation receives an ExcelContext with access to the Excel app and workbook.
    /// Use this ONLY when the operation performs async I/O (file operations, etc.).
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="operation">Async operation to execute</param>
    /// <param name="timeout">Optional timeout override. If not specified, uses default (2 minutes). Maximum is 5 minutes.</param>
    /// <param name="cancellationToken">Optional cancellation token</param>
    /// <returns>Result of the operation</returns>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Excel COM error occurred</exception>
    /// <exception cref="TimeoutException">Operation exceeded the timeout period</exception>
    Task<T> ExecuteAsync<T>(
        Func<ExcelContext, CancellationToken, Task<T>> operation,
        TimeSpan? timeout = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Saves changes to the workbook.
    /// This is an explicit save - changes are NOT automatically saved on dispose.
    /// </summary>
    /// <param name="timeout">Optional timeout override. If not specified, uses default (2 minutes). Maximum is 5 minutes.</param>
    /// <param name="cancellationToken">Optional cancellation token</param>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Save failed (e.g., file is read-only)</exception>
    /// <exception cref="TimeoutException">Save operation exceeded the timeout period</exception>
    Task SaveAsync(TimeSpan? timeout = null, CancellationToken cancellationToken = default);

}
