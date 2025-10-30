using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table Data Model integration operations (AddToDataModel)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> AddToDataModelAsync(IExcelBatch batch, string tableName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "add-to-data-model" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? modelTables = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Check if workbook has a Data Model (Model object)
                dynamic? model = null;
                try
                {
                    model = ctx.Book.Model;
                    if (model == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = "Workbook does not have a Data Model. Data Model is only available in Excel 2013+ with Power Pivot enabled.";
                        return result;
                    }
                }
                catch
                {
                    result.Success = false;
                    result.ErrorMessage = "Data Model not available. Ensure Excel has Power Pivot add-in enabled.";
                    return result;
                }

                // Check if table is already in the Data Model
                try
                {
                    modelTables = model.ModelTables;
                    for (int i = 1; i <= modelTables.Count; i++)
                    {
                        dynamic? modelTable = null;
                        try
                        {
                            modelTable = modelTables.Item(i);
                            string sourceTableName = modelTable.SourceName;
                            if (sourceTableName == tableName || sourceTableName.EndsWith($"[{tableName}]"))
                            {
                                result.Success = false;
                                result.ErrorMessage = $"Table '{tableName}' is already in the Data Model";
                                return result;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref modelTable);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                }

                // Create a connection for the table
                string connectionName = $"WorkbookConnection_{tableName}";
                string connectionString = $"WORKSHEET;{ctx.Book.FullName}";
                string commandText = $"SELECT * FROM [{tableName}]";

                // Add table to Data Model using the correct Microsoft approach
                // Key insight: Use Connections.Add2() with CreateModelConnection=true
                // This automatically adds the table to the Data Model
                try
                {
                    dynamic workbookConnections = ctx.Book.Connections;

                    // Check if a connection for this table already exists
                    bool connectionExists = false;
                    for (int i = 1; i <= workbookConnections.Count; i++)
                    {
                        dynamic? existingConn = null;
                        try
                        {
                            existingConn = workbookConnections.Item(i);
                            if (existingConn.Name == connectionName)
                            {
                                connectionExists = true;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref existingConn);
                        }
                    }

                    if (!connectionExists)
                    {
                        // Use Connections.Add2() with CreateModelConnection=true
                        // This is the documented approach for adding tables to Data Model
                        dynamic? newConnection = workbookConnections.Add2(
                            Name: connectionName,
                            Description: $"Excel Table: {tableName}",
                            ConnectionString: connectionString,
                            CommandText: commandText,
                            lCmdtype: 4, // xlCmdTable = 4 for Excel tables
                            CreateModelConnection: true, // KEY: This adds table to Data Model
                            ImportRelationships: false
                        );

                        if (newConnection != null)
                        {
                            ComUtilities.Release(ref newConnection);
                        }
                    }

                    if (workbookConnections != null)
                    {
                        ComUtilities.Release(ref workbookConnections!);
                    }
                }
                catch (Exception ex)
                {
                    // Fallback: Try the table's publish method if available
                    try
                    {
                        // Some Excel versions support Publish method on ListObject
                        table.Publish(null, false); // Publish to Data Model
                    }
                    catch (Exception publishEx)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Failed to add table to Data Model. " +
                                            $"Connections.Add2 failed: {ex.Message}. " +
                                            $"Table.Publish failed: {publishEx.Message}. " +
                                            $"Ensure Power Pivot is enabled and the Data Model is available.";
                        return result;
                    }
                }

                ComUtilities.Release(ref model);

                result.Success = true;
                result.SuggestedNextActions.Add("Use 'dm-list-tables' to verify the table is in the Data Model");
                result.SuggestedNextActions.Add($"Use 'dm-create-measure' to add DAX measures based on '{tableName}'");
                result.SuggestedNextActions.Add("Use 'dm-refresh' to refresh the Data Model");
                result.WorkflowHint = $"Table '{tableName}' added to Power Pivot Data Model.";

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
                // Only release variables that were actually used in this method
                if (modelTables != null)
                {
                    ComUtilities.Release(ref modelTables);
                }
                if (table != null)
                {
                    ComUtilities.Release(ref table);
                }
            }
        });
    }
#pragma warning restore CS1998
}
