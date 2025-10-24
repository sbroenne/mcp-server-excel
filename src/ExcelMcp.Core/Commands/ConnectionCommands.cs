using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.ComInterop;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Session;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection management commands - Core data layer (no console output)
/// Provides CRUD operations for Excel data connections (OLEDB, ODBC, Text, Web, etc.)
/// </summary>
public class ConnectionCommands : IConnectionCommands
{
    /// <summary>
    /// Lists all connections in a workbook
    /// </summary>
    public ConnectionListResult List(string filePath)
    {
        var result = new ConnectionListResult { FilePath = filePath };

        return ExcelSession.Execute(filePath, save: false, (excel, workbook) =>
        {
            try
            {
                dynamic connections = workbook.Connections;

                for (int i = 1; i <= connections.Count; i++)
                {
                    try
                    {
                        dynamic conn = connections.Item(i);

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
        });
    }

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    public ConnectionViewResult View(string filePath, string connectionName)
    {
        var result = new ConnectionViewResult
        {
            FilePath = filePath,
            ConnectionName = connectionName
        };

        return ExcelSession.Execute(filePath, save: false, (excel, workbook) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

                if (conn == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' not found";
                    return result;
                }

                result.Type = ConnectionHelpers.GetConnectionTypeName(conn.Type);
                result.IsPowerQuery = PowerQueryHelpers.IsPowerQueryConnection(conn);

                // Get connection string (sanitized for security)
                string? rawConnectionString = GetConnectionString(conn);
                result.ConnectionString = ConnectionHelpers.SanitizeConnectionString(rawConnectionString);

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

                result.DefinitionJson = JsonSerializer.Serialize(definition, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

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
    /// Imports connection from JSON file
    /// </summary>
    public OperationResult Import(string filePath, string connectionName, string jsonFilePath)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
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

            return ExcelSession.Execute(filePath, save: true, (excel, workbook) =>
            {
                try
                {
                    // Check if connection already exists
                    dynamic? existing = ComUtilities.FindConnection(workbook, connectionName);
                    if (existing != null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Connection '{connectionName}' already exists. Use 'update' to modify existing connection.";
                        return result;
                    }

                    // Create new connection based on type
                    CreateConnection(workbook, connectionName, definition);

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
    /// Exports connection to JSON file
    /// </summary>
    public OperationResult Export(string filePath, string connectionName, string jsonFilePath)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "export"
        };

        return ExcelSession.Execute(filePath, save: false, (excel, workbook) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

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
                    result.ErrorMessage = $"Connection '{connectionName}' is a Power Query connection. Use 'pq-export' command instead.";
                    return result;
                }

                // Build connection definition
                var definition = new ConnectionDefinition
                {
                    Name = connectionName,
                    Description = conn.Description?.ToString() ?? "",
                    Type = ConnectionHelpers.GetConnectionTypeName(conn.Type),
                    ConnectionString = ConnectionHelpers.SanitizeConnectionString(GetConnectionString(conn)),
                    CommandText = GetCommandText(conn),
                    CommandType = GetCommandType(conn),
                    BackgroundQuery = GetBackgroundQuerySetting(conn),
                    RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn),
                    SavePassword = false, // Never export with SavePassword = true (security)
                    RefreshPeriod = GetRefreshPeriod(conn)
                };

                // Serialize to JSON
                string json = JsonSerializer.Serialize(definition, new JsonSerializerOptions
                {
                    WriteIndented = true
                });

                // Write to file
                File.WriteAllText(jsonFilePath, json);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error exporting connection: {ex.Message}";
                return result;
            }
        });
    }

    /// <summary>
    /// Updates existing connection from JSON file
    /// </summary>
    public OperationResult Update(string filePath, string connectionName, string jsonFilePath)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
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

            return ExcelSession.Execute(filePath, save: true, (excel, workbook) =>
            {
                try
                {
                    dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

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
    public OperationResult Refresh(string filePath, string connectionName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "refresh"
        };

        return ExcelSession.Execute(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

                if (conn == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' not found";
                    return result;
                }

                // Refresh the connection
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
        });
    }

    /// <summary>
    /// Deletes a connection
    /// </summary>
    public OperationResult Delete(string filePath, string connectionName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "delete"
        };

        return ExcelSession.Execute(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

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
                PowerQueryHelpers.RemoveQueryTables(workbook, connectionName);

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

    /// <summary>
    /// Loads connection data to a worksheet
    /// </summary>
    public OperationResult LoadTo(string filePath, string connectionName, string sheetName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "loadto"
        };

        return ExcelSession.Execute(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

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
                    result.ErrorMessage = $"Connection '{connectionName}' is a Power Query connection. Use 'pq-loadto' command instead.";
                    return result;
                }

                // Find or create target sheet
                dynamic sheets = workbook.Worksheets;
                dynamic? targetSheet = null;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic sheet = sheets.Item(i);
                    if (sheet.Name.ToString().Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetSheet = sheet;
                        break;
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // Remove existing QueryTables first
                PowerQueryHelpers.RemoveQueryTables(workbook, connectionName);

                // Create QueryTable to load data
                var options = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = connectionName,
                    RefreshImmediately = true
                };

                CreateQueryTableForConnection(targetSheet, connectionName, conn, options);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error loading connection to sheet: {ex.Message}";
                return result;
            }
        });
    }

    /// <summary>
    /// Gets connection properties
    /// </summary>
    public ConnectionPropertiesResult GetProperties(string filePath, string connectionName)
    {
        var result = new ConnectionPropertiesResult
        {
            FilePath = filePath,
            ConnectionName = connectionName
        };

        return ExcelSession.Execute(filePath, save: false, (excel, workbook) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

                if (conn == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' not found";
                    return result;
                }

                result.BackgroundQuery = GetBackgroundQuerySetting(conn);
                result.RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn);
                result.SavePassword = GetSavePasswordSetting(conn);
                result.RefreshPeriod = GetRefreshPeriod(conn);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error getting connection properties: {ex.Message}";
                return result;
            }
        });
    }

    /// <summary>
    /// Sets connection properties
    /// </summary>
    public OperationResult SetProperties(string filePath, string connectionName,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "set-properties"
        };

        return ExcelSession.Execute(filePath, save: true, (excel, workbook) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

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
                    result.ErrorMessage = $"Connection '{connectionName}' is a Power Query connection. Power Query properties cannot be modified directly.";
                    return result;
                }

                // Update properties if specified
                SetConnectionProperty(conn, "BackgroundQuery", backgroundQuery);
                SetConnectionProperty(conn, "RefreshOnFileOpen", refreshOnFileOpen);
                SetConnectionProperty(conn, "SavePassword", savePassword);
                SetConnectionProperty(conn, "RefreshPeriod", refreshPeriod);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting connection properties: {ex.Message}";
                return result;
            }
        });
    }

    /// <summary>
    /// Tests connection without refreshing data
    /// </summary>
    public OperationResult Test(string filePath, string connectionName)
    {
        var result = new OperationResult
        {
            FilePath = filePath,
            Action = "test"
        };

        return ExcelSession.Execute(filePath, save: false, (excel, workbook) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(workbook, connectionName);

                if (conn == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' not found";
                    return result;
                }

                // Get connection type
                int connType = conn.Type;
                string typeName = ConnectionHelpers.GetConnectionTypeName(connType);

                // For Text (4) and Web (5) connections, connection string might not be accessible
                // until a QueryTable is created. Just verify the connection object exists.
                if (connType == 4 || connType == 5)
                {
                    result.Success = true;
                    return result;
                }

                // For other connection types (OLEDB, ODBC), validate connection string
                string? connectionString = GetConnectionString(conn);

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    result.Success = false;
                    result.ErrorMessage = "Connection has no connection string configured";
                    return result;
                }

                // Connection exists and is accessible
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Connection test failed: {ex.Message}";
                return result;
            }
        });
    }

    #region Helper Methods

    private static bool GetBackgroundQuerySetting(dynamic conn)
    {
        try
        {
            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                return conn.OLEDBConnection?.BackgroundQuery ?? false;
            }
            else if (connType == 2) // ODBC
            {
                return conn.ODBCConnection?.BackgroundQuery ?? false;
            }
            else if (connType == 3) // Text
            {
                return conn.TextConnection?.BackgroundQuery ?? false;
            }
            else if (connType == 4) // Web
            {
                return conn.WebConnection?.BackgroundQuery ?? false;
            }
        }
        catch
        {
            // Property not available
        }

        return false;
    }

    private static bool GetRefreshOnFileOpenSetting(dynamic conn)
    {
        try
        {
            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                return conn.OLEDBConnection?.RefreshOnFileOpen ?? false;
            }
            else if (connType == 2) // ODBC
            {
                return conn.ODBCConnection?.RefreshOnFileOpen ?? false;
            }
            else if (connType == 3) // Text
            {
                return conn.TextConnection?.RefreshOnFileOpen ?? false;
            }
            else if (connType == 4) // Web
            {
                return conn.WebConnection?.RefreshOnFileOpen ?? false;
            }
        }
        catch
        {
            // Property not available
        }

        return false;
    }

    private static bool GetSavePasswordSetting(dynamic conn)
    {
        try
        {
            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                return conn.OLEDBConnection?.SavePassword ?? false;
            }
            else if (connType == 2) // ODBC
            {
                return conn.ODBCConnection?.SavePassword ?? false;
            }
            else if (connType == 3) // Text
            {
                return conn.TextConnection?.SavePassword ?? false;
            }
            else if (connType == 4) // Web
            {
                return conn.WebConnection?.SavePassword ?? false;
            }
        }
        catch
        {
            // Property not available
        }

        return false;
    }

    private static int GetRefreshPeriod(dynamic conn)
    {
        try
        {
            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                return conn.OLEDBConnection?.RefreshPeriod ?? 0;
            }
            else if (connType == 2) // ODBC
            {
                return conn.ODBCConnection?.RefreshPeriod ?? 0;
            }
        }
        catch
        {
            // Property not available
        }

        return 0;
    }

    private static DateTime? GetLastRefreshDate(dynamic conn)
    {
        try
        {
            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                var refreshDate = conn.OLEDBConnection?.RefreshDate;
                if (refreshDate != null)
                {
                    return refreshDate;
                }
            }
            else if (connType == 2) // ODBC
            {
                var refreshDate = conn.ODBCConnection?.RefreshDate;
                if (refreshDate != null)
                {
                    return refreshDate;
                }
            }
        }
        catch
        {
            // Property not available
        }

        return null;
    }

    private static string? GetConnectionString(dynamic conn)
    {
        try
        {
            int connType = conn.Type;
            string? connectionString = null;

            if (connType == 1) // OLEDB
            {
                connectionString = conn.OLEDBConnection?.Connection?.ToString();
            }
            else if (connType == 2) // ODBC
            {
                connectionString = conn.ODBCConnection?.Connection?.ToString();
            }
            else if (connType == 4) // TEXT (xlConnectionTypeTEXT)
            {
                // Try to get from TextConnection first
                dynamic textConn = conn.TextConnection;
                if (textConn != null)
                {
                    try { connectionString = textConn.Connection?.ToString(); } catch { }
                }
            }
            else if (connType == 5) // WEB (xlConnectionTypeWEB)
            {
                // Try to get from WebConnection first
                dynamic webConn = conn.WebConnection;
                if (webConn != null)
                {
                    try { connectionString = webConn.Connection?.ToString(); } catch { }
                }
            }

            // If we still don't have a connection string, try the root ConnectionString property
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                try
                {
                    connectionString = conn.ConnectionString?.ToString();
                }
                catch
                {
                    // Property not available
                }
            }

            return connectionString;
        }
        catch
        {
            // Property not available
        }

        return null;
    }

    private static string? GetCommandText(dynamic conn)
    {
        try
        {
            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                return conn.OLEDBConnection?.CommandText?.ToString();
            }
            else if (connType == 2) // ODBC
            {
                return conn.ODBCConnection?.CommandText?.ToString();
            }
            else if (connType == 3) // Text
            {
                return conn.TextConnection?.CommandText?.ToString();
            }
            else if (connType == 4) // Web
            {
                return conn.WebConnection?.CommandText?.ToString();
            }
        }
        catch
        {
            // Property not available
        }

        return null;
    }

    private static string? GetCommandType(dynamic conn)
    {
        try
        {
            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                int? cmdType = conn.OLEDBConnection?.CommandType;
                return cmdType switch
                {
                    1 => "Cube",
                    2 => "SQL",
                    3 => "Table",
                    4 => "Default",
                    5 => "List",
                    _ => cmdType?.ToString()
                };
            }
            else if (connType == 2) // ODBC
            {
                int? cmdType = conn.ODBCConnection?.CommandType;
                return cmdType switch
                {
                    1 => "Cube",
                    2 => "SQL",
                    3 => "Table",
                    4 => "Default",
                    5 => "List",
                    _ => cmdType?.ToString()
                };
            }
        }
        catch
        {
            // Property not available
        }

        return null;
    }

    private static object GetConnectionProperties(dynamic conn)
    {
        return new
        {
            BackgroundQuery = GetBackgroundQuerySetting(conn),
            RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn),
            SavePassword = GetSavePasswordSetting(conn),
            RefreshPeriod = GetRefreshPeriod(conn),
            LastRefresh = GetLastRefreshDate(conn)
        };
    }

    private static void CreateConnection(dynamic workbook, string connectionName, ConnectionDefinition definition)
    {
        // Validate required fields
        if (string.IsNullOrWhiteSpace(definition.ConnectionString))
        {
            throw new InvalidOperationException("ConnectionString is required to create a connection.");
        }

        try
        {
            dynamic connections = workbook.Connections;

            // Create connection using Connections.Add() method
            // Per Microsoft documentation: https://learn.microsoft.com/en-us/office/vba/api/excel.connections.add
            // Parameters: Name (Required), Description (Required), ConnectionString (Required),
            //             CommandText (Required), lCmdtype (Optional), CreateModelConnection (Optional), ImportRelationships (Optional)
            dynamic newConnection = connections.Add(
                Name: connectionName,
                Description: definition.Description ?? "",
                ConnectionString: definition.ConnectionString,
                CommandText: definition.CommandText ?? ""
                // Note: Omitting optional parameters (lCmdtype, CreateModelConnection, ImportRelationships)
                // to let Excel use defaults
            );

            // Connection created successfully
            // Note: Setting additional properties after creation can cause COM errors
            // If needed in the future, handle carefully based on connection type
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to create connection '{connectionName}': {ex.Message}", ex);
        }
    }

    private static void UpdateConnectionProperties(dynamic conn, ConnectionDefinition definition)
    {
        try
        {
            // Update description
            if (!string.IsNullOrWhiteSpace(definition.Description))
            {
                conn.Description = definition.Description;
            }

            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                var oledb = conn.OLEDBConnection;
                if (oledb != null)
                {
                    if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                    {
                        oledb.Connection = definition.ConnectionString;
                    }
                    if (!string.IsNullOrWhiteSpace(definition.CommandText))
                    {
                        oledb.CommandText = definition.CommandText;
                    }
                    if (definition.BackgroundQuery.HasValue)
                    {
                        oledb.BackgroundQuery = definition.BackgroundQuery.Value;
                    }
                    if (definition.RefreshOnFileOpen.HasValue)
                    {
                        oledb.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                    }
                    if (definition.SavePassword.HasValue)
                    {
                        oledb.SavePassword = definition.SavePassword.Value;
                    }
                    if (definition.RefreshPeriod.HasValue)
                    {
                        oledb.RefreshPeriod = definition.RefreshPeriod.Value;
                    }
                }
            }
            else if (connType == 2) // ODBC
            {
                var odbc = conn.ODBCConnection;
                if (odbc != null)
                {
                    if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                    {
                        odbc.Connection = definition.ConnectionString;
                    }
                    if (!string.IsNullOrWhiteSpace(definition.CommandText))
                    {
                        odbc.CommandText = definition.CommandText;
                    }
                    if (definition.BackgroundQuery.HasValue)
                    {
                        odbc.BackgroundQuery = definition.BackgroundQuery.Value;
                    }
                    if (definition.RefreshOnFileOpen.HasValue)
                    {
                        odbc.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                    }
                    if (definition.SavePassword.HasValue)
                    {
                        odbc.SavePassword = definition.SavePassword.Value;
                    }
                    if (definition.RefreshPeriod.HasValue)
                    {
                        odbc.RefreshPeriod = definition.RefreshPeriod.Value;
                    }
                }
            }
            else if (connType == 3) // Text
            {
                var text = conn.TextConnection;
                if (text != null)
                {
                    if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                    {
                        text.Connection = definition.ConnectionString;
                    }
                    if (!string.IsNullOrWhiteSpace(definition.CommandText))
                    {
                        text.CommandText = definition.CommandText;
                    }
                    if (definition.BackgroundQuery.HasValue)
                    {
                        text.BackgroundQuery = definition.BackgroundQuery.Value;
                    }
                    if (definition.RefreshOnFileOpen.HasValue)
                    {
                        text.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                    }
                    if (definition.SavePassword.HasValue)
                    {
                        text.SavePassword = definition.SavePassword.Value;
                    }
                    if (definition.RefreshPeriod.HasValue)
                    {
                        text.RefreshPeriod = definition.RefreshPeriod.Value;
                    }
                }
            }
            else if (connType == 4) // Web
            {
                var web = conn.WebConnection;
                if (web != null)
                {
                    if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                    {
                        web.Connection = definition.ConnectionString;
                    }
                    if (!string.IsNullOrWhiteSpace(definition.CommandText))
                    {
                        web.CommandText = definition.CommandText;
                    }
                    if (definition.BackgroundQuery.HasValue)
                    {
                        web.BackgroundQuery = definition.BackgroundQuery.Value;
                    }
                    if (definition.RefreshOnFileOpen.HasValue)
                    {
                        web.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                    }
                    if (definition.SavePassword.HasValue)
                    {
                        web.SavePassword = definition.SavePassword.Value;
                    }
                    if (definition.RefreshPeriod.HasValue)
                    {
                        web.RefreshPeriod = definition.RefreshPeriod.Value;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to update connection properties: {ex.Message}", ex);
        }
    }

    private static void SetConnectionProperty<T>(dynamic conn, string propertyName, T? value) where T : struct
    {
        if (!value.HasValue) return;

        try
        {
            int connType = conn.Type;

            if (connType == 1) // OLEDB
            {
                var oledb = conn.OLEDBConnection;
                if (oledb != null)
                {
                    SetProperty(oledb, propertyName, value.Value);
                }
            }
            else if (connType == 2) // ODBC
            {
                var odbc = conn.ODBCConnection;
                if (odbc != null)
                {
                    SetProperty(odbc, propertyName, value.Value);
                }
            }
            else if (connType == 3) // Text
            {
                var text = conn.TextConnection;
                if (text != null)
                {
                    SetProperty(text, propertyName, value.Value);
                }
            }
            else if (connType == 4) // Web
            {
                var web = conn.WebConnection;
                if (web != null)
                {
                    SetProperty(web, propertyName, value.Value);
                }
            }
        }
        catch
        {
            // Property not available for this connection type
        }
    }

    private static void SetProperty<T>(dynamic obj, string propertyName, T value)
    {
        try
        {
            // Use reflection to set property dynamically
            var type = obj.GetType();
            var property = type.GetProperty(propertyName);
            if (property != null && property.CanWrite)
            {
                property.SetValue(obj, value);
            }
        }
        catch
        {
            // Property doesn't exist or can't be set
        }
    }

    private static void CreateQueryTableForConnection(dynamic targetSheet, string connectionName,
        dynamic conn, PowerQueryHelpers.QueryTableOptions options)
    {
        // For regular connections (not Power Query), we need connection string
        string? connectionString = GetConnectionString(conn);
        string? commandText = GetCommandText(conn);

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            throw new InvalidOperationException("Connection has no connection string");
        }

        // Command text can be empty for some connection types (Text, Web)
        // Use empty string if not provided
        if (string.IsNullOrWhiteSpace(commandText))
        {
            commandText = "";
        }

        dynamic queryTables = targetSheet.QueryTables;
        dynamic queryTable = queryTables.Add(connectionString, targetSheet.Range["A1"], commandText);

        queryTable.Name = options.Name.Replace(" ", "_");
        queryTable.RefreshStyle = 1; // xlInsertDeleteCells
        queryTable.BackgroundQuery = options.BackgroundQuery;
        queryTable.RefreshOnFileOpen = options.RefreshOnFileOpen;
        queryTable.SavePassword = options.SavePassword;
        queryTable.PreserveColumnInfo = options.PreserveColumnInfo;
        queryTable.PreserveFormatting = options.PreserveFormatting;
        queryTable.AdjustColumnWidth = options.AdjustColumnWidth;

        if (options.RefreshImmediately)
        {
            queryTable.Refresh(false);
        }
    }

    #endregion
}

/// <summary>
/// Connection definition for JSON import/export
/// </summary>
internal class ConnectionDefinition
{
    public string Name { get; set; } = "";
    public string? Description { get; set; }
    public string Type { get; set; } = "";
    public string? ConnectionString { get; set; }
    public string? CommandText { get; set; }
    public string? CommandType { get; set; }
    public bool? BackgroundQuery { get; set; }
    public bool? RefreshOnFileOpen { get; set; }
    public bool? SavePassword { get; set; }
    public int? RefreshPeriod { get; set; }
}
