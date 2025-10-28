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
            dynamic? connections = null;
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

                // Check if connection already exists
                connections = ctx.Book.Connections;
                bool connectionExists = false;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic? conn = null;
                    try
                    {
                        conn = connections.Item(i);
                        if (conn.Name == connectionName)
                        {
                            connectionExists = true;
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref conn);
                    }
                }

                // Add table to Data Model
                // Using numeric constant for xlConnectionTypeOLEDB = 3
                if (!connectionExists)
                {
                    try
                    {
                        dynamic? newConnection = connections.Add2(
                            connectionName,
                            "Connection to Excel Table",
                            connectionString,
                            commandText,
                            3, // xlConnectionTypeOLEDB
                            true, // SSO (not used for local)
                            false // AddToModel parameter
                        );
                        ComUtilities.Release(ref newConnection);
                    }
                    catch
                    {
                        // Connection might not be needed in some Excel versions
                        // Continue anyway
                    }
                }

                // Add the table to the model using ModelTables.Add
                try
                {
                    modelTables = model.ModelTables;
                    dynamic? newModelTable = modelTables.Add(
                        connectionName,
                        tableName
                    );
                    ComUtilities.Release(ref newModelTable);
                    ComUtilities.Release(ref modelTables);
                }
                catch (Exception ex)
                {
                    // Try alternative approach: use Publish to Data Model
                    try
                    {
                        // Some Excel versions support PublishToDataModel on ListObject
                        table.Publish(null, false); // Publish to Data Model
                    }
                    catch
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Failed to add table to Data Model: {ex.Message}. Ensure Power Pivot is enabled.";
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
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref modelTables);
                ComUtilities.Release(ref table);
            }
        });
    }
#pragma warning restore CS1998
}
