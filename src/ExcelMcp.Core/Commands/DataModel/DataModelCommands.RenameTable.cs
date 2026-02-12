using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Table Rename operation.
/// Implements COM-first approach with Power Query fallback for PQ-backed tables.
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public RenameResult RenameTable(IExcelBatch batch, string oldName, string newName)
    {
        return batch.Execute((ctx, _) =>
        {
            var result = new RenameResult
            {
                ObjectType = "data-model-table",
                OldName = oldName,
                NewName = newName,
                NormalizedOldName = RenameNameRules.Normalize(oldName),
                NormalizedNewName = RenameNameRules.Normalize(newName)
            };

            // Validate old name is not empty
            if (RenameNameRules.IsEmpty(result.NormalizedOldName))
            {
                result.Success = false;
                result.ErrorMessage = "Old table name cannot be empty or whitespace.";
                return result;
            }

            // Validate new name is not empty
            if (RenameNameRules.IsEmpty(result.NormalizedNewName))
            {
                result.Success = false;
                result.ErrorMessage = "New table name cannot be empty or whitespace.";
                return result;
            }

            // No-op when normalized names are exactly equal
            if (RenameNameRules.IsNoOp(result.NormalizedOldName, result.NormalizedNewName))
            {
                result.Success = true;
                return result;
            }

            // Check if workbook has Data Model
            if (!HasDataModelTables(ctx.Book))
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.NoDataModelTables();
                return result;
            }

            dynamic? model = null;
            dynamic? table = null;
            dynamic? sourceConnection = null;
            try
            {
                model = ctx.Book.Model;

                // Find target table (case-insensitive lookup per FindModelTable)
                table = FindModelTable(model, result.NormalizedOldName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.TableNotFound(result.NormalizedOldName);
                    return result;
                }

                // Collect existing table names for conflict detection
                var existingNames = new List<string>();
                ForEachTable(model, (Action<dynamic, int>)((t, _) =>
                {
                    existingNames.Add(ComUtilities.SafeGetString(t, "Name"));
                }));

                // Check for conflicts (case-insensitive, excluding target)
                if (RenameNameRules.HasConflict(existingNames, result.NormalizedNewName, result.NormalizedOldName))
                {
                    result.Success = false;
                    result.ErrorMessage = $"A table named '{result.NormalizedNewName}' already exists (case-insensitive match).";
                    return result;
                }

                // ModelTable.Name is read-only per Microsoft documentation.
                // Direct rename is not possible - must use Power Query rename for PQ-backed tables.
                // See: https://learn.microsoft.com/en-us/office/vba/excel/concepts/about-the-powerpivot-model-object-in-excel

                // Get the source connection to check if this is a PQ-backed table
                sourceConnection = table.SourceWorkbookConnection;
                if (sourceConnection == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Cannot rename table '{result.NormalizedOldName}': " +
                        "Direct rename is not supported and table has no source connection. " +
                        "Only Power Query-backed Data Model tables can be renamed.";
                    return result;
                }

                // Check if this is a Power Query connection
                if (!PowerQuery.PowerQueryHelpers.IsPowerQueryConnection(sourceConnection))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Cannot rename table '{result.NormalizedOldName}': " +
                        "Direct rename is not supported. Table is not backed by Power Query. " +
                        "Only Power Query-backed Data Model tables can be renamed.";
                    return result;
                }

                // Extract Power Query name from connection (format: "Query - {QueryName}")
                string connectionName = sourceConnection.Name?.ToString() ?? string.Empty;
                if (!connectionName.StartsWith("Query - ", StringComparison.OrdinalIgnoreCase))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Cannot rename table '{result.NormalizedOldName}': " +
                        "Power Query connection name format is unexpected: '{connectionName}'.";
                    return result;
                }

                string pqName = connectionName["Query - ".Length..];

                // Find and rename the underlying Power Query
                dynamic? targetQuery = null;
                dynamic? oleDbConnection = null;
                try
                {
                    targetQuery = ComUtilities.FindQuery(ctx.Book, pqName);
                    if (targetQuery == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Cannot rename table '{result.NormalizedOldName}': " +
                            $"Associated Power Query '{pqName}' not found.";
                        return result;
                    }

                    // Step 1: Rename the Power Query
                    targetQuery.Name = result.NormalizedNewName;

                    // Step 2: Update the connection name to match the new query name
                    // Connection name format: "Query - {QueryName}"
                    string newConnectionName = $"Query - {result.NormalizedNewName}";
                    sourceConnection.Name = newConnectionName;

                    // Step 3: Update the connection string to reference the new query name
                    // Connection string format: "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={QueryName}"
                    oleDbConnection = sourceConnection.OLEDBConnection;
                    if (oleDbConnection != null)
                    {
                        string? currentConnectionString = oleDbConnection.Connection?.ToString();
                        if (!string.IsNullOrEmpty(currentConnectionString))
                        {
                            // Replace Location={oldName} with Location={newName}
                            // Handle both exact match and partial match scenarios
                            string oldLocation = $"Location={pqName}";
                            string newLocation = $"Location={result.NormalizedNewName}";

                            if (currentConnectionString.Contains(oldLocation, StringComparison.OrdinalIgnoreCase))
                            {
                                string newConnectionString = currentConnectionString.Replace(
                                    oldLocation,
                                    newLocation,
                                    StringComparison.OrdinalIgnoreCase);
                                oleDbConnection.Connection = newConnectionString;
                            }
                        }
                    }

                    // Step 4: Refresh the Data Model to attempt table name update
                    // ModelTable.Name is read-only and cached from the connection at creation time.
                    // Refreshing the model DOES NOT update the table name - this is a known Excel limitation.
                    try
                    {
                        model.Refresh();
                    }
#pragma warning disable CA1031 // Catch more specific exception - Model.Refresh() can throw many COM exception types
                    catch (Exception)
                    {
                        // Model refresh may fail for various reasons (data source issues, etc.)
                        // This is best-effort and not critical to the operation
                    }
#pragma warning restore CA1031

                    // Step 5: Verify the table name actually updated using CASE-SENSITIVE comparison
                    // Excel's Data Model table names are immutable after creation.
                    // Even though we renamed the Power Query and connection, the ModelTable.Name
                    // remains cached at its original value. This is an Excel/COM API limitation.
                    //
                    // Note: FindModelTable uses case-insensitive lookup, so we must re-check the
                    // actual table name returned to confirm the rename truly succeeded.
                    ComUtilities.Release(ref table!);
                    table = FindModelTable(model, result.NormalizedNewName);

                    if (table != null)
                    {
                        // Table found - but verify the name matches EXACTLY (case-sensitive)
                        // FindModelTable uses case-insensitive lookup, so "testtable" would match "TestTable"
                        string actualName = ComUtilities.SafeGetString(table, "Name");
                        if (string.Equals(actualName, result.NormalizedNewName, StringComparison.Ordinal))
                        {
                            // Table name matches exactly - rename succeeded
                            result.Success = true;
                            return result;
                        }
                        // Table found but name doesn't match exactly - rename failed
                    }

                    // Table not found with new name - rename failed due to Excel limitation
                    // Rollback the Power Query and connection names to maintain consistency
                    try
                    {
                        targetQuery.Name = pqName;  // Restore original PQ name
                        sourceConnection.Name = connectionName;  // Restore original connection name
                        if (oleDbConnection != null)
                        {
                            string? currentConnectionString = oleDbConnection.Connection?.ToString();
                            if (!string.IsNullOrEmpty(currentConnectionString))
                            {
                                string newLocation = $"Location={result.NormalizedNewName}";
                                string oldLocation = $"Location={pqName}";
                                if (currentConnectionString.Contains(newLocation, StringComparison.OrdinalIgnoreCase))
                                {
                                    string restoredConnectionString = currentConnectionString.Replace(
                                        newLocation,
                                        oldLocation,
                                        StringComparison.OrdinalIgnoreCase);
                                    oleDbConnection.Connection = restoredConnectionString;
                                }
                            }
                        }
                    }
#pragma warning disable CA1031 // Catch more specific exception - Rollback is best-effort, must not throw
                    catch (Exception)
                    {
                        // Rollback failed - best effort cleanup, cannot propagate
                    }
#pragma warning restore CA1031

                    result.Success = false;
                    result.ErrorMessage = $"Cannot rename Data Model table '{result.NormalizedOldName}': " +
                        "Excel's Data Model table names are immutable after creation. " +
                        "The underlying Power Query and connection were temporarily renamed but have been rolled back. " +
                        "To rename a Data Model table, you must delete it and recreate it with the new name.";
                    return result;
                }
                finally
                {
                    ComUtilities.Release(ref oleDbConnection!);
                    ComUtilities.Release(ref targetQuery!);
                }
            }
            finally
            {
                ComUtilities.Release(ref sourceConnection!);
                ComUtilities.Release(ref table!);
                ComUtilities.Release(ref model!);
            }
        });
    }
}


