using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Refresh operations
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public OperationResult Refresh(IExcelBatch batch, string? tableName = null, TimeSpan? timeout = null)
    {
        var effectiveTimeout = timeout.HasValue && timeout.Value > TimeSpan.Zero
            ? timeout.Value
            : TimeSpan.FromMinutes(2);
        using var timeoutCts = new CancellationTokenSource(effectiveTimeout);

        try
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.Model? model = null;
                try
                {
                    // Check if workbook has Data Model
                    if (!HasDataModelTables(ctx.Book))
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                    }

                    model = ctx.Book.Model;

                    if (tableName != null)
                    {
                        // Refresh specific table
                        Excel.ModelTable? table = FindModelTable(model!, tableName);
                        if (table == null)
                        {
                            throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(tableName));
                        }

                        try
                        {
                            OleMessageFilter.EnterLongOperation();
                            try
                            {
                                table.Refresh();
                            }
                            finally
                            {
                                OleMessageFilter.ExitLongOperation();
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref table);
                        }
                    }
                    else
                    {
                        // Refresh entire model
                        OleMessageFilter.EnterLongOperation();
                        try
                        {
                            model.Refresh();
                        }
                        catch (Exception refreshEx)
                        {
                            // Model.Refresh() may not be supported in all Excel versions
                            // Fall back to refreshing tables individually
                            throw new InvalidOperationException($"Model-level refresh not supported. Try refreshing tables individually. Error: {refreshEx.Message}", refreshEx);
                        }
                        finally
                        {
                            OleMessageFilter.ExitLongOperation();
                        }
                    }

                    // NOTE: CUBEVALUE formulas may still show #N/A after refresh.
                    // Application.Calculate() and CalculateFull() can throw COM errors (0x800AC472).
                    // This is a known Excel COM limitation - CUBE functions require interactive Excel.
                    // See: https://github.com/sbroenne/mcp-server-excel/issues/313
                }
                finally
                {
                    ComUtilities.Release(ref model);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }, timeoutCts.Token);
        }
        catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
        {
            throw new TimeoutException(
                $"Data Model refresh timed out after {effectiveTimeout.TotalSeconds:F0} seconds for '{Path.GetFileName(batch.WorkbookPath)}'.");
        }
    }
}



