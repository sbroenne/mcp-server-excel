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
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Skip connections that have COM access issues
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
                throw new InvalidOperationException($"Connection '{connectionName}' not found.");
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
    public void Create(IExcelBatch batch, string connectionName,
        string connectionString, string? commandText = null, string? description = null)
    {
        batch.Execute((ctx, ct) =>
        {
            // Create connection definition
            var definition = new ConnectionDefinition
            {
                Name = connectionName,
                Description = description ?? "",
                ConnectionString = connectionString,
                CommandText = commandText ?? "",
                CommandType = string.IsNullOrWhiteSpace(commandText) ? null : "SQL",
                SavePassword = false // Default to secure setting
            };

            // Create the connection using existing helper method
            CreateConnection(ctx.Book, connectionName, definition);
            return 0;
        });
    }

    /// <summary>
    /// Refreshes connection data
    /// </summary>
    public void Refresh(IExcelBatch batch, string connectionName)
    {
        Refresh(batch, connectionName, timeout: null);
    }

    /// <summary>
    /// Refreshes connection data with timeout
    /// </summary>
    public void Refresh(IExcelBatch batch, string connectionName, TimeSpan? timeout)
    {
        var effectiveTimeout = timeout ?? TimeSpan.FromMinutes(5);
        using var timeoutCts = new CancellationTokenSource(effectiveTimeout);

        batch.Execute((ctx, ct) =>
        {
            dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found.");
            }

            // Check if this is a Power Query connection (handle separately)
            if (PowerQueryHelpers.IsPowerQueryConnection(conn))
            {
                // Check if this is an orphaned Power Query connection
                if (PowerQueryHelpers.IsOrphanedPowerQueryConnection(ctx.Book, conn))
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' is an orphaned Power Query connection with no corresponding query. Use connection 'delete' to remove it.");
                }
                throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Use powerquery 'refresh' instead.");
            }

            // Pure COM passthrough - just refresh the connection
            conn.Refresh();
            return 0;
        }, timeoutCts.Token);  // Extended timeout (default 5 minutes) for slow data sources
    }

    /// <summary>
    /// Deletes a connection
    /// </summary>
    public void Delete(IExcelBatch batch, string connectionName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found.");
            }

            // Check if this is a Power Query connection
            if (PowerQueryHelpers.IsPowerQueryConnection(conn))
            {
                // Check if this is an orphaned Power Query connection (no corresponding query exists)
                // Orphaned connections can be safely deleted via the connection API
                if (!PowerQueryHelpers.IsOrphanedPowerQueryConnection(ctx.Book, conn))
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Use powerquery with action 'Delete' instead.");
                }
                // Orphaned connection - allow deletion to proceed
            }

            // Remove associated QueryTables first
            PowerQueryHelpers.RemoveQueryTables(ctx.Book, connectionName);

            // Delete the connection
            conn.Delete();
            return 0;
        });
    }
}



