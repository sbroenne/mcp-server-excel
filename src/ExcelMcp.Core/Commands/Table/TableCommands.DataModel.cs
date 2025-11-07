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
    public async Task<OperationResult> AddToDataModelAsync(IExcelBatch batch, string tableName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "table-add-to-datamodel" };
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? model = null;
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

                // Data Model is always available in Excel 2013+ (no need to check)
                model = ctx.Book.Model;
                modelTables = model.ModelTables;

                // Check if table is already in the Data Model via ModelTables
                for (int i = 1; i <= modelTables.Count; i++)
                {
                    dynamic? modelTable = null;
                    try
                    {
                        modelTable = modelTables.Item(i);
                        string sourceTableName = modelTable.SourceName;
                        if (sourceTableName == tableName || sourceTableName.EndsWith($"[{tableName}]", StringComparison.Ordinal))
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

                // Create a connection for the table using the sigma_coding VBA pattern
                // ConnectionString: "WORKSHEET;{DirectoryPath}" (directory only, not full file path!)
                // CommandText: "{WorkbookName}!{TableName}" (not SQL query!)
                // lCmdtype: xlCmdExcel = 7 (THE KEY - not 4 or 8!)
                const int xlCmdExcel = 7;
                string connectionName = $"WorkbookConnection_{ctx.Book.Name}!{tableName}";

                // Add table to Data Model using sigma_coding VBA pattern
                dynamic? workbookConnections = null;
                dynamic? newConnection = null;
                try
                {
                    workbookConnections = ctx.Book.Connections;

                    // Double-check: Connection name shouldn't exist
                    for (int i = 1; i <= workbookConnections.Count; i++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = workbookConnections.Item(i);
                            if (conn.Name == connectionName)
                            {
                                result.Success = false;
                                result.ErrorMessage = $"Table '{tableName}' is already in the Data Model";
                                return result;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }

                    // Create the connection using EXACT pattern from sigma_coding VBA
                    newConnection = workbookConnections.Add2(
                        connectionName,                          // Name
                        $"Excel Table: {tableName}",             // Description
                        $"WORKSHEET;{ctx.Book.Path}",            // ConnectionString: "WORKSHEET;{DirectoryPath}"
                        $"{ctx.Book.Name}!{tableName}",          // CommandText: "{WorkbookName}!{TableName}"
                        xlCmdExcel,                              // lCmdtype: 7 (THE CRITICAL DIFFERENCE!)
                        true,                                    // CreateModelConnection: true
                        false                                    // ImportRelationships: false
                    );
                }
                catch (Exception ex)
                {
                    // Build diagnostic message
                    var diagnostics = new List<string> { $"Connections.Add2 exception: {ex.Message}" };
                    if (ex.InnerException != null)
                    {
                        diagnostics.Add($"Inner exception: {ex.InnerException.Message}");
                    }

                    result.Success = false;
                    result.ErrorMessage = $"Failed to add table to Data Model. {string.Join(" -- ", diagnostics)}";
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref newConnection);
                    ComUtilities.Release(ref workbookConnections);
                }

                // Table is immediately available in Data Model - no refresh needed
                // Connections.Add2() makes the table accessible for relationships/measures instantly

                result.Success = true;
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
                // Release COM objects
                ComUtilities.Release(ref modelTables);
                ComUtilities.Release(ref model);
                ComUtilities.Release(ref table);
            }
        });
    }
}
