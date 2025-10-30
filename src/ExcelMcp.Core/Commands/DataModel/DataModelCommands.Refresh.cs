using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Refresh operations
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(IExcelBatch batch, string? tableName = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = tableName != null ? $"model-refresh-table:{tableName}" : "model-refresh"
        };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModel();
                    return result;
                }

                model = ctx.Book.Model;

                if (tableName != null)
                {
                    // Refresh specific table
                    dynamic? table = ComUtilities.FindModelTable(model, tableName);
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
                        result.SuggestedNextActions =
                        [
                            $"Table '{tableName}' refreshed successfully",
                            "Use 'model-list-tables' to verify record counts"
                        ];
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
                        result.SuggestedNextActions =
                        [
                            "All Data Model tables refreshed successfully",
                            "Use 'model-list-tables' to verify record counts"
                        ];
                    }
                    catch (Exception refreshEx)
                    {
                        // Model.Refresh() may not be supported in all Excel versions
                        // Fall back to refreshing tables individually
                        result.ErrorMessage = $"Model-level refresh not supported. Try refreshing tables individually. Error: {refreshEx.Message}";
                        result.Success = false;
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
        });
    }
}
