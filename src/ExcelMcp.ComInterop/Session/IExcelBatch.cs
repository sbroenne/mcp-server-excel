namespace Sbroenne.ExcelMcp.ComInterop.Session;

using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Represents a batch of Excel operations that share a single Excel instance.
/// Implements IDisposable to ensure proper COM cleanup.
/// </summary>
/// <remarks>
/// Use this interface via ExcelSession.BeginBatch() for multi-operation workflows.
/// The batch keeps Excel and the workbook open until disposed, enabling efficient
/// execution of multiple commands without repeated Excel startup/shutdown overhead.
///
/// <para><b>Lifecycle:</b></para>
/// <list type="bullet">
/// <item>Created via ExcelSession.BeginBatch(filePath)</item>
/// <item>Operations executed via Execute()</item>
/// <item>Optional explicit save via Save()</item>
/// <item>Disposed via Dispose() or "using" pattern</item>
/// </list>
///
/// <para><b>Example:</b></para>
/// <code>
/// using var batch = ExcelSession.BeginBatch("workbook.xlsx");
///
/// // Execute COM operations
/// batch.Execute((ctx, ct) => {
///     ctx.Book.Worksheets.Add("Data");
///     return 0;
/// });
///
/// batch.Execute((ctx, ct) => {
///     ctx.Book.Worksheets["Data"].Range["A1"].Value = "Header";
///     return 0;
/// });
///
/// // Get content from Excel
/// var content = batch.Execute((ctx, ct) => {
///     return ctx.Book.Range["A1"].Formula;
/// });
///
/// // Explicit save
/// batch.Save();
/// </code>
/// </remarks>
public interface IExcelBatch : IDisposable
{
    /// <summary>
    /// Gets the path to the Excel workbook this batch operates on.
    /// For multi-workbook batches, this is the primary (first) workbook.
    /// </summary>
    string WorkbookPath { get; }

    /// <summary>
    /// Gets the logger instance for diagnostic output.
    /// Returns NullLogger if no logger was provided during construction.
    /// </summary>
    Microsoft.Extensions.Logging.ILogger Logger { get; }

    /// <summary>
    /// Gets all workbooks currently open in this batch, keyed by normalized file path.
    /// For single-workbook batches, contains one entry.
    /// For multi-workbook batches (cross-workbook operations), contains all open workbooks.
    /// </summary>
    IReadOnlyDictionary<string, Excel.Workbook> Workbooks { get; }

    /// <summary>
    /// Gets the COM Workbook object for a specific file path.
    /// </summary>
    /// <param name="filePath">Path to the workbook (will be normalized)</param>
    /// <returns>Excel.Workbook COM object</returns>
    /// <exception cref="KeyNotFoundException">Workbook not found in this batch</exception>
    Excel.Workbook GetWorkbook(string filePath);

    /// <summary>
    /// Executes a void COM operation within this batch.
    /// The operation receives an ExcelContext with access to the Excel app and workbook.
    /// Use this overload for void operations that don't need to return values.
    /// All Excel COM operations are synchronous - file I/O should be handled outside the batch.
    /// </summary>
    /// <param name="operation">COM operation to execute</param>
    /// <param name="cancellationToken">Optional cancellation token for caller-controlled cancellation</param>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Excel COM error occurred</exception>
    /// <exception cref="OperationCanceledException">Operation was cancelled via cancellationToken</exception>
    void Execute(
        Action<ExcelContext, CancellationToken> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Executes a COM operation within this batch.
    /// The operation receives an ExcelContext with access to the Excel app and workbook.
    /// All Excel COM operations are synchronous - file I/O should be handled outside the batch.
    /// </summary>
    /// <typeparam name="T">Return type of the operation</typeparam>
    /// <param name="operation">COM operation to execute</param>
    /// <param name="cancellationToken">Optional cancellation token for caller-controlled cancellation</param>
    /// <returns>Result of the operation</returns>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Excel COM error occurred</exception>
    /// <exception cref="OperationCanceledException">Operation was cancelled via cancellationToken</exception>
    T Execute<T>(
        Func<ExcelContext, CancellationToken, T> operation,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Saves changes to the workbook.
    /// This is an explicit save - changes are NOT automatically saved on dispose.
    /// </summary>
    /// <param name="cancellationToken">Optional cancellation token for caller-controlled cancellation</param>
    /// <exception cref="ObjectDisposedException">Batch has already been disposed</exception>
    /// <exception cref="InvalidOperationException">Save failed (e.g., file is read-only)</exception>
    /// <exception cref="OperationCanceledException">Save operation was cancelled via cancellationToken</exception>
    void Save(CancellationToken cancellationToken = default);

    /// <summary>
    /// Checks if the underlying Excel process is still alive.
    /// </summary>
    /// <returns>
    /// True if Excel process exists and hasn't exited.
    /// False if process has crashed, was killed, or process ID wasn't captured.
    /// </returns>
    /// <remarks>
    /// Use this to detect dead Excel processes before attempting operations.
    /// If this returns false, the session should be closed and recreated.
    /// </remarks>
    bool IsExcelProcessAlive();

    /// <summary>
    /// Gets the Excel process ID, if captured.
    /// </summary>
    /// <returns>Process ID, or null if not captured during startup.</returns>
    int? ExcelProcessId { get; }

    /// <summary>
    /// Gets the operation timeout for this batch.
    /// All Execute() calls will timeout after this duration.
    /// Default is 5 minutes (from ComInteropConstants.DefaultOperationTimeout).
    /// </summary>
    TimeSpan OperationTimeout { get; }

}




