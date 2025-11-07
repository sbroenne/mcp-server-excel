using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;

namespace Sbroenne.ExcelMcp.Core.Commands.QueryTable;

/// <summary>
/// QueryTable lifecycle operations (create, update)
/// </summary>
public partial class QueryTableCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> CreateFromConnectionAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string connectionName, string range = "A1",
        QueryTableCreateOptions? options = null)
    {
        return await batch.Execute((ctx, ct) =>
        {
            var result = new OperationResult
            {
                FilePath = batch.WorkbookPath,
                Action = "create-from-connection"
            };

            // Find the connection
            dynamic? connection = ComUtilities.FindConnection(ctx.Book, connectionName);
            if (connection == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Connection '{connectionName}' not found";
                return result;
            }

            dynamic? worksheet = null;
            dynamic? queryTables = null;
            dynamic? queryTable = null;
            dynamic? targetRange = null;

            try
            {
                // Get target worksheet
                try
                {
                    worksheet = ctx.Book.Worksheets.Item(sheetName);
                }
                catch
                {
                    result.Success = false;
                    result.ErrorMessage = $"Worksheet '{sheetName}' not found";
                    return result;
                }

                queryTables = worksheet.QueryTables;
                targetRange = worksheet.Range[range];

                // Get connection string and command text based on connection type
                int connType = connection.Type;
                string connectionString = "";
                string commandText = "";

                if (connType == 1) // OLEDB
                {
                    dynamic? oledb = connection.OLEDBConnection;
                    try
                    {
                        connectionString = oledb?.Connection?.ToString() ?? "";
                        commandText = oledb?.CommandText?.ToString() ?? "";
                    }
                    finally
                    {
                        ComUtilities.Release(ref oledb);
                    }
                }
                else if (connType == 2) // ODBC
                {
                    dynamic? odbc = connection.ODBCConnection;
                    try
                    {
                        connectionString = odbc?.Connection?.ToString() ?? "";
                        commandText = odbc?.CommandText?.ToString() ?? "";
                    }
                    finally
                    {
                        ComUtilities.Release(ref odbc);
                    }
                }
                else if (connType == 3 || connType == 4) // TEXT (3) or WEB (4)
                {
                    // Try TextConnection first, fall back to WebConnection
                    try
                    {
                        dynamic? text = connection.TextConnection;
                        try
                        {
                            connectionString = text?.Connection?.ToString() ?? "";
                        }
                        finally
                        {
                            ComUtilities.Release(ref text);
                        }
                    }
                    catch
                    {
                        dynamic? web = connection.WebConnection;
                        try
                        {
                            connectionString = web?.Connection?.ToString() ?? "";
                        }
                        finally
                        {
                            ComUtilities.Release(ref web);
                        }
                    }
                }

                // Create QueryTable
                queryTable = queryTables.Add(connectionString, targetRange, commandText);
                queryTable.Name = queryTableName.Replace(" ", "_");

                // Apply options
                options ??= new QueryTableCreateOptions();
                ApplyQueryTableOptions(queryTable, options);

                // Refresh immediately if requested (default true)
                if (options.RefreshImmediately)
                {
                    queryTable.Refresh(false);  // CRITICAL: Synchronous for persistence
                }

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to create QueryTable: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref targetRange);
                ComUtilities.Release(ref queryTable);
                ComUtilities.Release(ref queryTables);
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref connection);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateFromQueryAsync(IExcelBatch batch, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        QueryTableCreateOptions? options = null)
    {
        return await batch.Execute((ctx, ct) =>
        {
            var result = new OperationResult
            {
                FilePath = batch.WorkbookPath,
                Action = "create-from-query"
            };

            // Verify the Power Query exists
            dynamic? query = ComUtilities.FindQuery(ctx.Book, queryName);
            if (query == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Power Query '{queryName}' not found";
                return result;
            }
            ComUtilities.Release(ref query);

            dynamic? worksheet = null;
            try
            {
                // Get target worksheet
                try
                {
                    worksheet = ctx.Book.Worksheets.Item(sheetName);
                }
                catch
                {
                    result.Success = false;
                    result.ErrorMessage = $"Worksheet '{sheetName}' not found";
                    return result;
                }

                // Use existing PowerQueryHelpers infrastructure
                var queryTableOptions = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = queryTableName,
                    BackgroundQuery = options?.BackgroundQuery ?? false,
                    RefreshOnFileOpen = options?.RefreshOnFileOpen ?? false,
                    SavePassword = options?.SavePassword ?? false,
                    PreserveColumnInfo = options?.PreserveColumnInfo ?? true,
                    PreserveFormatting = options?.PreserveFormatting ?? true,
                    AdjustColumnWidth = options?.AdjustColumnWidth ?? true,
                    RefreshImmediately = options?.RefreshImmediately ?? true
                };

                PowerQueryHelpers.CreateQueryTable(worksheet, queryName, queryTableOptions);

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to create QueryTable from Power Query: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref worksheet);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdatePropertiesAsync(IExcelBatch batch, string queryTableName,
        QueryTableUpdateOptions options)
    {
        return await batch.Execute((ctx, ct) =>
        {
            var result = new OperationResult
            {
                FilePath = batch.WorkbookPath,
                Action = "update-properties"
            };

            dynamic? queryTable = ComUtilities.FindQueryTable(ctx.Book, queryTableName);
            if (queryTable == null)
            {
                result.Success = false;
                result.ErrorMessage = $"QueryTable '{queryTableName}' not found";
                return result;
            }

            try
            {
                // Update only the properties that are specified
                if (options.BackgroundQuery.HasValue)
                    queryTable.BackgroundQuery = options.BackgroundQuery.Value;

                if (options.RefreshOnFileOpen.HasValue)
                    queryTable.RefreshOnFileOpen = options.RefreshOnFileOpen.Value;

                if (options.SavePassword.HasValue)
                    queryTable.SavePassword = options.SavePassword.Value;

                if (options.PreserveColumnInfo.HasValue)
                    queryTable.PreserveColumnInfo = options.PreserveColumnInfo.Value;

                if (options.PreserveFormatting.HasValue)
                    queryTable.PreserveFormatting = options.PreserveFormatting.Value;

                if (options.AdjustColumnWidth.HasValue)
                    queryTable.AdjustColumnWidth = options.AdjustColumnWidth.Value;

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to update QueryTable properties: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }

            return result;
        });
    }

    /// <summary>
    /// Applies QueryTableCreateOptions to a QueryTable COM object
    /// </summary>
    private static void ApplyQueryTableOptions(dynamic queryTable, QueryTableCreateOptions options)
    {
        try
        {
            queryTable.BackgroundQuery = options.BackgroundQuery;
            queryTable.RefreshOnFileOpen = options.RefreshOnFileOpen;
            queryTable.SavePassword = options.SavePassword;
            queryTable.PreserveColumnInfo = options.PreserveColumnInfo;
            queryTable.PreserveFormatting = options.PreserveFormatting;
            queryTable.AdjustColumnWidth = options.AdjustColumnWidth;

            // Apply refresh style for cell insertion behavior
            queryTable.RefreshStyle = 1; // xlInsertDeleteCells
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to apply QueryTable options: {ex.Message}", ex);
        }
    }
}
