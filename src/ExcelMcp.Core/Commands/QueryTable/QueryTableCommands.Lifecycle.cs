using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;

namespace Sbroenne.ExcelMcp.Core.Commands.QueryTable;

/// <summary>
/// QueryTable lifecycle operations (create, update)
/// </summary>
public partial class QueryTableCommands
{
    /// <inheritdoc />
    public OperationResult CreateFromConnection(IExcelBatch batch, string sheetName,
        string queryTableName, string connectionName, string range = "A1",
        PowerQueryHelpers.QueryTableCreateOptions? options = null)
    {
        return batch.Execute((ctx, ct) =>
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

                // Get connection string and command text using ConnectionHelpers
                string? connectionString = ConnectionHelpers.GetConnectionString(connection);
                string? commandText = ConnectionHelpers.GetCommandText(connection);

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' has no connection string";
                    return result;
                }

                // Use unified QueryTable creation method
                options ??= new PowerQueryHelpers.QueryTableCreateOptions { Name = queryTableName };
                var createOptions = new PowerQueryHelpers.QueryTableCreateOptions
                {
                    Name = queryTableName,
                    Range = range,
                    ConnectionString = connectionString,
                    CommandText = commandText ?? "",
                    BackgroundQuery = options.BackgroundQuery,
                    RefreshOnFileOpen = options.RefreshOnFileOpen,
                    SavePassword = options.SavePassword,
                    PreserveColumnInfo = options.PreserveColumnInfo,
                    PreserveFormatting = options.PreserveFormatting,
                    AdjustColumnWidth = options.AdjustColumnWidth,
                    RefreshImmediately = options.RefreshImmediately
                };

                PowerQueryHelpers.CreateQueryTable(worksheet, createOptions);

                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to create QueryTable: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref connection);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public OperationResult CreateFromQuery(IExcelBatch batch, string sheetName,
        string queryTableName, string queryName, string range = "A1",
        PowerQueryHelpers.QueryTableCreateOptions? options = null)
    {
        return batch.Execute((ctx, ct) =>
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

                // Use unified QueryTable creation method
                options ??= new PowerQueryHelpers.QueryTableCreateOptions { Name = queryTableName };
                var createOptions = new PowerQueryHelpers.QueryTableCreateOptions
                {
                    Name = queryTableName,
                    Range = range,
                    BackgroundQuery = options.BackgroundQuery,
                    RefreshOnFileOpen = options.RefreshOnFileOpen,
                    SavePassword = options.SavePassword,
                    PreserveColumnInfo = options.PreserveColumnInfo,
                    PreserveFormatting = options.PreserveFormatting,
                    AdjustColumnWidth = options.AdjustColumnWidth,
                    RefreshImmediately = options.RefreshImmediately
                };

                PowerQueryHelpers.CreateQueryTable(worksheet, createOptions, queryName);

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
    public OperationResult UpdateProperties(IExcelBatch batch, string queryTableName,
        QueryTableUpdateOptions options)
    {
        return batch.Execute((ctx, ct) =>
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
}

