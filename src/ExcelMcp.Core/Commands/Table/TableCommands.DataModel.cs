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

                // Add table to Data Model using the correct Microsoft approach:
                // 1. Create a legacy WorkbookConnection for the table
                // 2. Use Model.AddConnection() to add it to the Data Model
                // Per Microsoft docs: Model.AddConnection only works with legacy (non-model) connections
                
                string connectionName = $"WorkbookConnection_{tableName}";
                string connectionString = $"WORKSHEET;{ctx.Book.FullName}";
                string commandText = tableName; // For Excel tables, use table name directly
                
                dynamic? workbookConnections = null;
                dynamic? legacyConnection = null;
                try
                {
                    workbookConnections = ctx.Book.Connections;

                    // Check if a legacy connection for this table already exists
                    bool connectionExists = false;
                    for (int i = 1; i <= workbookConnections.Count; i++)
                    {
                        dynamic? existingConn = null;
                        try
                        {
                            existingConn = workbookConnections.Item(i);
                            if (existingConn.Name == connectionName)
                            {
                                legacyConnection = existingConn;
                                connectionExists = true;
                                break;
                            }
                        }
                        finally
                        {
                            if (!connectionExists && existingConn != null)
                            {
                                ComUtilities.Release(ref existingConn);
                            }
                        }
                    }

                    if (!connectionExists)
                    {
                        // Create legacy connection (NOT a model connection)
                        // Use Add2 with CreateModelConnection=false to create legacy connection first
                        legacyConnection = workbookConnections.Add2(
                            Name: connectionName,
                            Description: $"Excel Table: {tableName}",
                            ConnectionString: connectionString,
                            CommandText: commandText,
                            lCmdtype: 2, // xlCmdDefault = 2
                            CreateModelConnection: false, // Create as legacy connection first
                            ImportRelationships: false
                        );
                    }

                    // Now add the legacy connection to the Data Model
                    // This is the documented Microsoft approach: Model.AddConnection(legacyConnection)
                    dynamic? modelConnection = model.AddConnection(legacyConnection);
                    
                    if (modelConnection != null)
                    {
                        ComUtilities.Release(ref modelConnection);
                    }
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Failed to add table to Data Model: {ex.Message}. " +
                                        $"Ensure Power Pivot is enabled and the Data Model is available.";
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref legacyConnection);
                    ComUtilities.Release(ref workbookConnections);
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
