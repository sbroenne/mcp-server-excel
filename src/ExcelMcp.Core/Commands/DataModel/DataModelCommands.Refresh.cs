using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Refresh operations
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(IExcelBatch batch, string? tableName = null)
    {
        return await RefreshAsync(batch, tableName, TimeSpan.FromMinutes(2));  // Default 2 minutes for Data Model refresh, LLM can override
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(IExcelBatch batch, string? tableName, TimeSpan? timeout)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = tableName != null ? $"model-refresh-table:{tableName}" : "model-refresh"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                    return result;
                }

                model = ctx.Book.Model;

                if (tableName != null)
                {
                    // Refresh specific table
                    dynamic? table = FindModelTable(model, tableName);
                    if (table == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = DataModelErrorMessages.TableNotFound(tableName);
                        return result;
                    }

                    try
                    {
                        table.Refresh();
                        result.Success = true;
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
                        result.Success = true;
                    }
                    catch (Exception refreshEx)
                    {
                        result.Success = false;
                        // Model.Refresh() may not be supported in all Excel versions
                        // Fall back to refreshing tables individually
                        result.ErrorMessage = $"Model-level refresh not supported. Try refreshing tables individually. Error: {refreshEx.Message}";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("refreshing Data Model", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        }, timeout: timeout ?? TimeSpan.FromMinutes(2));  // Default 2 minutes for Data Model refresh, LLM can override
    }
}
