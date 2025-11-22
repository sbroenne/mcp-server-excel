using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection lifecycle operations (List, View, Import, Export, Update, Delete)
/// </summary>
public partial class ConnectionCommands
{
    private static readonly JsonSerializerOptions s_jsonOptions = new() { WriteIndented = true };

    /// <summary>
    /// Lists all connections in a workbook
    /// </summary>
    public ConnectionListResult List(IExcelBatch batch)
    {
        var result = new ConnectionListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? connections = null;

            try
            {
                connections = ctx.Book.Connections;

                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic? conn = null;
                    try
                    {
                        conn = connections.Item(i);

                        var connInfo = new ConnectionInfo
                        {
                            Name = conn.Name?.ToString() ?? "",
                            Description = conn.Description?.ToString() ?? "",
                            Type = ConnectionHelpers.GetConnectionTypeName(conn.Type),
                            IsPowerQuery = PowerQueryHelpers.IsPowerQueryConnection(conn),
                            BackgroundQuery = GetBackgroundQuerySetting(conn),
                            RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn),
                            LastRefresh = GetLastRefreshDate(conn)
                        };

                        result.Connections.Add(connInfo);
                    }
                    catch
                    {
                        // Skip connections that can't be read
                        continue;
                    }
                    finally
                    {
                        ComUtilities.Release(ref conn);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref connections);
            }
        });
    }

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    public ConnectionViewResult View(IExcelBatch batch, string connectionName)
    {
        var result = new ConnectionViewResult
        {
            FilePath = batch.WorkbookPath,
            ConnectionName = connectionName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Connection '{connectionName}' not found";
                return result;
            }

            result.Type = ConnectionHelpers.GetConnectionTypeName(conn.Type);
            result.IsPowerQuery = PowerQueryHelpers.IsPowerQueryConnection(conn);

            // Get connection string (raw for LLM usage - sanitization removed)
            string? rawConnectionString = GetConnectionString(conn);
            result.ConnectionString = rawConnectionString ?? "";

            // Get command text and type
            result.CommandText = GetCommandText(conn);
            result.CommandType = GetCommandType(conn);

            // Build comprehensive JSON definition
            var definition = new
            {
                Name = connectionName,
                Type = result.Type,
                Description = conn.Description?.ToString() ?? "",
                IsPowerQuery = result.IsPowerQuery,
                ConnectionString = result.ConnectionString,
                CommandText = result.CommandText,
                CommandType = result.CommandType,
                Properties = GetConnectionProperties(conn)
            };

            result.DefinitionJson = JsonSerializer.Serialize(definition, s_jsonOptions);

            result.Success = true;
            return result;
        });
    }

    /// <summary>
    /// Creates a new connection in the workbook
    /// </summary>
    public OperationResult Create(IExcelBatch batch, string connectionName,
        string connectionString, string? commandText = null, string? description = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "create"
        };

        return batch.Execute((ctx, ct) =>
        {
            // Create connection definition
            var definition = new ConnectionDefinition
            {
                Name = connectionName,
                Description = description ?? "",
                ConnectionString = connectionString,
                CommandText = commandText ?? "",
                SavePassword = false // Default to secure setting
            };

            // Create the connection using existing helper method
            CreateConnection(ctx.Book, connectionName, definition);

            result.Success = true;
            return result;
        });
    }

    /// <summary>
    /// Refreshes connection data
    /// </summary>
    public OperationResult Refresh(IExcelBatch batch, string connectionName)
    {
        return Refresh(batch, connectionName, timeout: null);
    }

    /// <summary>
    /// Refreshes connection data with timeout
    /// </summary>
    public OperationResult Refresh(IExcelBatch batch, string connectionName, TimeSpan? timeout)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "refresh"
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Connection '{connectionName}' not found";
                return result;
            }

            // Check if this is a Power Query connection (handle separately)
            if (PowerQueryHelpers.IsPowerQueryConnection(conn))
            {
                result.Success = false;
                result.ErrorMessage = $"Connection '{connectionName}' is a Power Query connection. Use excel_powerquery 'refresh' instead.";
                return result;
            }

            // Pure COM passthrough - just refresh the connection
            conn.Refresh();

            result.Success = true;
            return result;
        });  // Default 2 minutes for connection refresh, LLM can override
    }

    /// <summary>
    /// Deletes a connection
    /// </summary>
    public OperationResult Delete(IExcelBatch batch, string connectionName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "delete"
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found");
            }

            // Check if this is a Power Query connection
            if (PowerQueryHelpers.IsPowerQueryConnection(conn))
            {
                throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Use 'pq-delete' command instead.");
            }

            // Remove associated QueryTables first
            PowerQueryHelpers.RemoveQueryTables(ctx.Book, connectionName);

            // Delete the connection
            conn.Delete();

            result.Success = true;
            return result;
        });
    }
}

