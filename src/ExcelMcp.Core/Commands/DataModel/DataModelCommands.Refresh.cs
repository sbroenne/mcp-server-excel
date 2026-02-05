using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Refresh operations
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public void Refresh(IExcelBatch batch, string? tableName = null, TimeSpan? timeout = null)
    {
        // timeout parameter reserved for future use (e.g., cancellation token support)
        _ = timeout;

        batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
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
                    dynamic? table = FindModelTable(model, tableName);
                    if (table == null)
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(tableName));
                    }

                    try
                    {
                        table.Refresh();
                    }
                    finally
                    {
                        ComUtilities.Release(ref table);
                    }
                }
                else
                {
                    // Refresh entire model
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

            return 0;
        });
    }
}



