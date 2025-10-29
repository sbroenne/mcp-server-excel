using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table column management operations (NEW)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> AddColumnAsync(IExcelBatch batch, string tableName, string columnName, int? position = null)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "add-column" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? newColumn = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Check if column already exists
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? column = null;
                    try
                    {
                        column = listColumns.Item(i);
                        if (column.Name == columnName)
                        {
                            result.Success = false;
                            result.ErrorMessage = $"Column '{columnName}' already exists in table '{tableName}'";
                            return result;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref column);
                    }
                }

                // Add column at specified position or at the end
                if (position.HasValue)
                {
                    newColumn = listColumns.Add(Position: position.Value);
                }
                else
                {
                    newColumn = listColumns.Add(); // Adds at end
                }

                newColumn.Name = columnName;

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table-info {tableName}' to verify column added");
                result.SuggestedNextActions.Add($"Use 'range-set-values' to populate new column data");
                result.WorkflowHint = $"Column '{columnName}' added to table '{tableName}'.";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref newColumn);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> RemoveColumnAsync(IExcelBatch batch, string tableName, string columnName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "remove-column" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Find column
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? col = null;
                    try
                    {
                        col = listColumns.Item(i);
                        if (col.Name == columnName)
                        {
                            column = col;
                            break;
                        }
                    }
                    finally
                    {
                        if (col != null && col.Name != columnName)
                        {
                            ComUtilities.Release(ref col);
                        }
                    }
                }

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{columnName}' not found in table '{tableName}'";
                    return result;
                }

                // Delete column
                column.Delete();

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table-info {tableName}' to verify column removed");
                result.WorkflowHint = $"Column '{columnName}' removed from table '{tableName}'.";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> RenameColumnAsync(IExcelBatch batch, string tableName, string oldName, string newName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "rename-column" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Find column
                listColumns = table.ListColumns;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? col = null;
                    try
                    {
                        col = listColumns.Item(i);
                        if (col.Name == oldName)
                        {
                            column = col;
                            break;
                        }
                    }
                    finally
                    {
                        if (col != null && col.Name != oldName)
                        {
                            ComUtilities.Release(ref col);
                        }
                    }
                }

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{oldName}' not found in table '{tableName}'";
                    return result;
                }

                // Check if new name already exists
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? col = null;
                    try
                    {
                        col = listColumns.Item(i);
                        if (col.Name == newName)
                        {
                            result.Success = false;
                            result.ErrorMessage = $"Column '{newName}' already exists in table '{tableName}'";
                            return result;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref col);
                    }
                }

                // Rename column
                column.Name = newName;

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table-info {tableName}' to verify column renamed");
                result.SuggestedNextActions.Add("Update any formulas using structured references with the old column name");
                result.WorkflowHint = $"Column renamed from '{oldName}' to '{newName}' in table '{tableName}'.";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }
#pragma warning restore CS1998
}
