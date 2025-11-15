using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Security;

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
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error listing connections: {ex.Message}";
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
            try
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
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error viewing connection: {ex.Message}";
                return result;
            }
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
            try
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
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error creating connection: {ex.Message}";
                return result;
            }
        });
    }

    /// <summary>
    /// Imports connection from JSON file
    /// </summary>
    public OperationResult Import(IExcelBatch batch, string connectionName, string jsonFilePath)
    {
        // Validate file path to prevent path traversal attacks
        jsonFilePath = PathValidator.ValidateExistingFile(jsonFilePath, nameof(jsonFilePath));

        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "import"
        };

        try
        {
            // Read and parse JSON definition
            if (!File.Exists(jsonFilePath))
            {
                result.Success = false;
                result.ErrorMessage = $"JSON file not found: {jsonFilePath}";
                return result;
            }

            string jsonContent = File.ReadAllText(jsonFilePath);
            var definition = JsonSerializer.Deserialize<ConnectionDefinition>(jsonContent);

            if (definition == null)
            {
                result.Success = false;
                result.ErrorMessage = "Failed to parse JSON connection definition";
                return result;
            }

            return batch.Execute((ctx, ct) =>
            {
                try
                {
                    // Check if connection already exists
                    dynamic? existing = ComUtilities.FindConnection(ctx.Book, connectionName);
                    if (existing != null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Connection '{connectionName}' already exists. Use 'update' to modify existing connection.";
                        return result;
                    }

                    // Create new connection based on type
                    CreateConnection(ctx.Book, connectionName, definition);

                    result.Success = true;
                    return result;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error importing connection: {ex.Message}";
                    return result;
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error reading JSON file: {ex.Message}";
            return result;
        }
    }

    /// <summary>
    /// Updates existing connection from JSON file
    /// </summary>
    public OperationResult UpdateProperties(IExcelBatch batch, string connectionName, string jsonFilePath)
    {
        // Validate file path to prevent path traversal attacks
        jsonFilePath = PathValidator.ValidateExistingFile(jsonFilePath, nameof(jsonFilePath));

        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "update"
        };

        try
        {
            // Read and parse JSON definition
            if (!File.Exists(jsonFilePath))
            {
                result.Success = false;
                result.ErrorMessage = $"JSON file not found: {jsonFilePath}";
                return result;
            }

            string jsonContent = File.ReadAllText(jsonFilePath);
            var definition = JsonSerializer.Deserialize<ConnectionDefinition>(jsonContent);

            if (definition == null)
            {
                result.Success = false;
                result.ErrorMessage = "Failed to parse JSON connection definition";
                return result;
            }

            return batch.Execute((ctx, ct) =>
            {
                try
                {
                    dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

                    if (conn == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Connection '{connectionName}' not found. Use 'import' to create new connection.";
                        return result;
                    }

                    // Check if this is a Power Query connection
                    if (PowerQueryHelpers.IsPowerQueryConnection(conn))
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Connection '{connectionName}' is a Power Query connection. Use 'pq-update' command instead.";
                        return result;
                    }

                    // Update connection properties
                    UpdateConnectionProperties(conn, definition);

                    result.Success = true;
                    return result;
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Error updating connection: {ex.Message}";
                    return result;
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error reading JSON file: {ex.Message}";
            return result;
        }
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
            try
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
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error refreshing connection: {ex.Message}";
                return result;
            }
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
            try
            {
                dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

                if (conn == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' not found";
                    return result;
                }

                // Check if this is a Power Query connection
                if (PowerQueryHelpers.IsPowerQueryConnection(conn))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' is a Power Query connection. Use 'pq-delete' command instead.";
                    return result;
                }

                // Remove associated QueryTables first
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, connectionName);

                // Delete the connection
                conn.Delete();

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error deleting connection: {ex.Message}";
                return result;
            }
        });
    }
}

