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
    public OperationResult Refresh(IExcelBatch batch, string? tableName = null)
    {
        return Refresh(batch, tableName, TimeSpan.FromMinutes(2));  // Default 2 minutes for Data Model refresh, LLM can override
    }

    /// <inheritdoc />
    public OperationResult Refresh(IExcelBatch batch, string? tableName, TimeSpan? timeout)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = tableName != null ? $"model-refresh-table:{tableName}" : "model-refresh"
        };

        return batch.Execute((ctx, ct) =>
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
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });  // Default 2 minutes for Data Model refresh, LLM can override
    }
}

