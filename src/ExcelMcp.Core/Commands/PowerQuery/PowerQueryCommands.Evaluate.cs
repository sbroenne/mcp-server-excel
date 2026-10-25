using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query evaluate operation - executes M code and returns results without creating a permanent query
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public PowerQueryEvaluateResult Evaluate(IExcelBatch batch, string mCode)
    {
        var result = new PowerQueryEvaluateResult
        {
            FilePath = batch.WorkbookPath,
            MCode = mCode
        };

        // Validate M code
        if (string.IsNullOrWhiteSpace(mCode))
        {
            throw new ArgumentException("M code is required for evaluate action", nameof(mCode));
        }

        // Generate unique temporary names
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var tempQueryName = $"__pq_eval_{uniqueId}";
        var tempSheetName = $"__pq_eval_{uniqueId}";

        using var timeoutCts = new CancellationTokenSource(batch.OperationTimeout);

        PowerQueryEvaluateResult Execute() => batch.Execute(EvaluateTemporaryQuery, timeoutCts.Token);

        [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
            Justification = "Retains the evaluation failure and every temporary-object cleanup failure, releases COM references, then throws the original or aggregate error; no failed evaluation returns success.")]
        PowerQueryEvaluateResult EvaluateTemporaryQuery(ExcelContext ctx, CancellationToken ct)
        {
            Excel.Queries? queriesCollection = null;
            Excel.WorkbookQuery? query = null;
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? tempSheet = null;
            Excel.ListObjects? listObjects = null;
            Excel.ListObject? listObject = null;
            Excel.QueryTable? queryTable = null;
            Excel.Range? range = null;
            Exception? operationError = null;

            try
            {
                ct.ThrowIfCancellationRequested();

                // STEP 1: Create temporary query with the M code
                queriesCollection = ctx.Book.Queries;
                query = queriesCollection.Add(tempQueryName, mCode);

                // STEP 2: Create temporary worksheet and load data to it
                worksheets = ctx.Book.Worksheets;
                tempSheet = (Excel.Worksheet)worksheets.Add();
                tempSheet.Name = tempSheetName;

                // STEP 3: Load query to worksheet using QueryTable
                // Connection string for Power Query - must include Extended Properties
                string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={tempQueryName};Extended Properties=\"\"";
                range = tempSheet.Range["A1"];

                listObjects = tempSheet.ListObjects;
                listObject = listObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcExternal,
                    connectionString,   // Source: connection string
                    Type.Missing,       // LinkSource
                    Excel.XlYesNoGuess.xlYes,
                    range               // Destination: starting cell
                );

                // Get the QueryTable to refresh (this executes the M code)
                queryTable = listObject.QueryTable;

                // Configure QueryTable to select from the query
                queryTable.CommandType = Excel.XlCmdType.xlCmdSql;
                queryTable.CommandText = $"SELECT * FROM [{tempQueryName}]";
                queryTable.BackgroundQuery = false; // Synchronous

                // STEP 4: Refresh to execute the M code (errors will throw via QueryTable.Refresh)
                // This path uses the same synchronous worksheet refresh mechanism as PowerQuery refresh/load.
                // Do NOT use EnterLongOperation here: rejecting inbound COM callbacks can deadlock the
                // MashupHost → Excel → STA callback chain that completes the synchronous refresh.
                try
                {
                    QueryTableRefreshHelper.RefreshSynchronously(queryTable, ct, "Power Query evaluation");
                }
                catch (COMException ex) when (TryWrapPowerQueryException(ex, out var queryError))
                {
                    throw queryError!;
                }

                ct.ThrowIfCancellationRequested();

                // STEP 5: Read the results from the worksheet
                // Get the data range from the ListObject
                Excel.Range? dataBodyRange = null;
                Excel.Range? headerRowRange = null;
                try
                {
                    // Get column names from header row
                    headerRowRange = listObject.HeaderRowRange;
                    if (headerRowRange != null)
                    {
                        object? headerValues = headerRowRange.Value2;
                        if (headerValues != null)
                        {
                            if (headerValues is object[,] headers2D)
                            {
                                for (int col = 1; col <= headers2D.GetLength(1); col++)
                                {
                                    ct.ThrowIfCancellationRequested();
                                    result.Columns.Add(headers2D[1, col]?.ToString() ?? $"Column{col}");
                                }
                            }
                            else if (headerValues is object[] headers1D)
                            {
                                for (int i = 0; i < headers1D.Length; i++)
                                {
                                    ct.ThrowIfCancellationRequested();
                                    result.Columns.Add(headers1D[i]?.ToString() ?? $"Column{i + 1}");
                                }
                            }
                            else
                            {
                                // Single cell
                                result.Columns.Add(headerValues?.ToString() ?? "Column1");
                            }
                        }
                    }
                    result.ColumnCount = result.Columns.Count;

                    // Get data rows
                    dataBodyRange = listObject.DataBodyRange;
                    if (dataBodyRange != null)
                    {
                        object? dataValues = dataBodyRange.Value2;
                        if (dataValues != null)
                        {
                            if (dataValues is object[,] data2D)
                            {
                                int rowCount = data2D.GetLength(0);
                                int colCount = data2D.GetLength(1);

                                for (int row = 1; row <= rowCount; row++)
                                {
                                    ct.ThrowIfCancellationRequested();
                                    var rowData = new List<object?>();
                                    for (int col = 1; col <= colCount; col++)
                                    {
                                        rowData.Add(ConvertCellValue(data2D[row, col]));
                                    }
                                    result.Rows.Add(rowData);
                                }
                            }
                            else if (dataValues is object[] data1D)
                            {
                                // Single row
                                var rowData = new List<object?>();
                                foreach (var val in data1D)
                                {
                                    ct.ThrowIfCancellationRequested();
                                    rowData.Add(ConvertCellValue(val));
                                }
                                result.Rows.Add(rowData);
                            }
                            else
                            {
                                // Single cell
                                result.Rows.Add([ConvertCellValue(dataValues)]);
                            }
                        }
                    }
                    result.RowCount = result.Rows.Count;
                }
                finally
                {
                    ComUtilities.Release(ref headerRowRange);
                    ComUtilities.Release(ref dataBodyRange);
                }

            }
            catch (Exception ex)
            {
                operationError = ex;
            }

            var cleanupErrors = new List<Exception>();
            try
            {
                if (listObject != null)
                {
                    listObject.Delete();
                }
            }
            catch (COMException ex)
            {
                cleanupErrors.Add(new InvalidOperationException(
                    $"Failed to delete temporary evaluate table for '{tempQueryName}'.",
                    ex));
            }

            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref listObject);
            ComUtilities.Release(ref listObjects);
            ComUtilities.Release(ref range);

            try
            {
                if (tempSheet != null)
                {
                    // Suppress alerts to avoid "Are you sure you want to delete?" prompt
                    Excel.Application? app = null;
                    try
                    {
                        app = ctx.Book.Application;
                        bool originalAlerts = app.DisplayAlerts;
                        app.DisplayAlerts = false;
                        try
                        {
                            tempSheet.Delete();
                        }
                        finally
                        {
                            app.DisplayAlerts = originalAlerts;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref app);
                    }
                }
            }
            catch (COMException ex)
            {
                cleanupErrors.Add(new InvalidOperationException(
                    $"Failed to delete temporary evaluate worksheet '{tempSheetName}'.",
                    ex));
            }

            ComUtilities.Release(ref tempSheet);
            ComUtilities.Release(ref worksheets);

            try
            {
                PowerQuery.PowerQueryHelpers.RemoveConnectionsByMashupLocation(
                    ctx.Book,
                    tempQueryName);
            }
            catch (Exception ex) when (ex is COMException or InvalidOperationException)
            {
                cleanupErrors.Add(new InvalidOperationException(
                    $"Failed to delete temporary evaluate connections for '{tempQueryName}'.",
                    ex));
            }

            try
            {
                if (query != null)
                {
                    query.Delete();
                }
            }
            catch (COMException ex)
            {
                cleanupErrors.Add(new InvalidOperationException(
                    $"Failed to delete temporary evaluate query '{tempQueryName}'.",
                    ex));
            }

            ComUtilities.Release(ref query);
            ComUtilities.Release(ref queriesCollection);

            if (cleanupErrors.Count > 0)
            {
                if (operationError != null)
                {
                    cleanupErrors.Insert(0, operationError);
                }

                var cleanupFailure = new PowerQueryCommandException(
                    $"Power Query evaluate cleanup failed for temporary identity '{tempQueryName}'. " +
                    "Inspect powerquery list and connection list before saving, or close without saving and retry.",
                    "Cleanup",
                    new AggregateException(cleanupErrors));
                if (operationError is OperationCanceledException canceled)
                    throw new OperationCanceledException(cleanupFailure.Message, cleanupFailure, canceled.CancellationToken);
                throw cleanupFailure;
            }

            if (operationError != null)
            {
                ExceptionDispatchInfo.Capture(operationError).Throw();
            }

            result.Success = true;
            return result;
        }

        try
        {
            return Execute();
        }
        catch (OperationCanceledException ex) when (timeoutCts.IsCancellationRequested)
        {
            throw new TimeoutException(
                $"Power Query evaluate timed out after {batch.OperationTimeout.TotalSeconds} seconds.",
                ex);
        }
    }

    /// <summary>
    /// Converts Excel cell values to JSON-friendly types
    /// </summary>
    private static object? ConvertCellValue(object? value)
    {
        if (value == null || value == DBNull.Value)
            return null;

        return value switch
        {
            DateTime dt => dt.ToString("O"), // ISO 8601 format
            double d => d,
            int i => i,
            long l => l,
            bool b => b,
            string s => s,
            _ => value.ToString()
        };
    }
}
