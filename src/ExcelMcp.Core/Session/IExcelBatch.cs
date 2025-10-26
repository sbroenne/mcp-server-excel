namespace Sbroenne.ExcelMcp.Core.Session;

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
/// // Multiple operations without reopening Excel
/// await batch.ExecuteAsync(async (ctx, ct) => {
///     ctx.Book.Worksheets.Add("Data");
///     return 0;
/// });
///
/// await batch.ExecuteAsync(async (ctx, ct) => {
///     ctx.Book.Worksheets["Data"].Range["A1"].Value = "Header";
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
    /// Executes an async operation within this batch.
    /// The operation receives an ExcelContext with access to the Excel app and workbook.
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="operation">Async operation to execute</param>
    /// <param name="cancellationToken">Optional cancellation token</param>
    /// <returns>Result of the operation</returns>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Excel COM error occurred</exception>
    Task<T> ExecuteAsync<T>(
        Func<ExcelContext, CancellationToken, ValueTask<T>> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Saves changes to the workbook.
    /// This is an explicit save - changes are NOT automatically saved on dispose.
    /// </summary>
    /// <param name="cancellationToken">Optional cancellation token</param>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Save failed (e.g., file is read-only)</exception>
    Task SaveAsync(CancellationToken cancellationToken = default);

}
