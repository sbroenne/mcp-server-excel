using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.PowerQuery;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection management commands - Core data layer (no console output)
/// Provides CRUD operations for Excel data connections (OLEDB, ODBC, Text, Web, etc.)
/// </summary>
public partial class ConnectionCommands : IConnectionCommands
{
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
            else if (connType is 3 or 4) // TEXT (type 3) or WEB (type 4) - Excel may report CSV as either
            {
                // Try TextConnection first, fall back to WebConnection
                try
                {
                    return conn.TextConnection?.BackgroundQuery ?? false;
                }
                catch
                {
                    try
                    {
                        return conn.WebConnection?.BackgroundQuery ?? false;
                    }
                    catch
                    {
                        return false;
                    }
                }
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
            else if (connType is 3 or 4) // TEXT (type 3) or WEB (type 4) - Excel may report CSV as either
            {
                // Try TextConnection first, fall back to WebConnection
                try
                {
                    return conn.TextConnection?.RefreshOnFileOpen ?? false;
                }
                catch
                {
                    try
                    {
                        return conn.WebConnection?.RefreshOnFileOpen ?? false;
                    }
                    catch
                    {
                        return false;
                    }
                }
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
            else if (connType is 3 or 4) // TEXT (type 3) or WEB (type 4) - Excel may report CSV as either
            {
                // Try TextConnection first, fall back to WebConnection
                try
                {
                    return conn.TextConnection?.SavePassword ?? false;
                }
                catch
                {
                    try
                    {
                        return conn.WebConnection?.SavePassword ?? false;
                    }
                    catch
                    {
                        return false;
                    }
                }
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
            else if (connType is 3 or 4) // TEXT (type 3) or WEB (type 4) - Excel may report CSV as either
            {
                // Try TextConnection first, fall back to WebConnection
                try
                {
                    return conn.TextConnection?.RefreshPeriod ?? 0;
                }
                catch
                {
                    try
                    {
                        return conn.WebConnection?.RefreshPeriod ?? 0;
                    }
                    catch
                    {
                        return 0;
                    }
                }
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

        dynamic? connections = null;
        dynamic? newConnection = null;

        try
        {
            connections = workbook.Connections;

            // Create connection using Connections.Add() method
            // Per Microsoft documentation: https://learn.microsoft.com/en-us/office/vba/api/excel.connections.add
            // Parameters: Name (Required), Description (Required), ConnectionString (Required),
            //             CommandText (Required), lCmdtype (Optional), CreateModelConnection (Optional), ImportRelationships (Optional)
            newConnection = connections.Add(
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
        finally
        {
            ComUtilities.Release(ref newConnection);
            ComUtilities.Release(ref connections);
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
            else if (connType is 3 or 4) // TEXT (type 3) or WEB (type 4) - Excel may report CSV files as either
            {
                // Excel has type 3/4 confusion: CSV files created with "TEXT;filepath" may be reported as type 4 (WEB)
                // Try TextConnection first (correct for type 3), fall back to WebConnection if that fails
                dynamic? textOrWeb = null!;
                try
                {
                    textOrWeb = conn.TextConnection; // Try TEXT first
                }
                catch
                {
                    try
                    {
                        textOrWeb = conn.WebConnection; // Fall back to WEB
                    }
                    catch
                    {
                        // Neither works - skip property updates
                    }
                }

                if (textOrWeb != null)
                {
                    if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                    {
                        textOrWeb.Connection = definition.ConnectionString;
                    }
                    if (!string.IsNullOrWhiteSpace(definition.CommandText))
                    {
                        textOrWeb.CommandText = definition.CommandText;
                    }
                    if (definition.BackgroundQuery.HasValue)
                    {
                        textOrWeb.BackgroundQuery = definition.BackgroundQuery.Value;
                    }
                    if (definition.RefreshOnFileOpen.HasValue)
                    {
                        textOrWeb.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                    }
                    if (definition.SavePassword.HasValue)
                    {
                        textOrWeb.SavePassword = definition.SavePassword.Value;
                    }
                    if (definition.RefreshPeriod.HasValue)
                    {
                        textOrWeb.RefreshPeriod = definition.RefreshPeriod.Value;
                    }
                }
            }
            else if (connType == 5) // XMLMAP (moved from 4 due to type 3/4 merge)
            {
                // XMLMAP connection properties - future implementation
                // For now, just update basic properties like description (already done above)
            }
            else if (connType == 6) // DATAFEED
            {
                // DATAFEED connection properties - future implementation
            }
            else if (connType == 7) // MODEL
            {
                // MODEL connection properties - future implementation
            }
            else if (connType == 8) // WORKSHEET
            {
                // WORKSHEET connection properties - future implementation
            }
            else if (connType == 9) // NOSOURCE
            {
                // NOSOURCE connection properties - future implementation
            }
            else
            {
                // Unknown connection type - skip property updates
                // Description was already updated above if provided
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
            else if (connType is 3 or 4) // TEXT (type 3) or WEB (type 4) - Excel may report CSV files as either
            {
                // Try TextConnection first, fall back to WebConnection
                dynamic? textOrWeb = null;
                try
                {
                    textOrWeb = conn.TextConnection;
                }
                catch
                {
                    try
                    {
                        textOrWeb = conn.WebConnection;
                    }
                    catch
                    {
                        // Neither works
                    }
                }

                if (textOrWeb != null)
                {
                    SetProperty(textOrWeb, propertyName, value.Value);
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

    private static void CreateQueryTableForConnection(
        dynamic targetSheet,
        dynamic conn,
        PowerQueryHelpers.QueryTableCreateOptions options)
    {
        // Get connection string and command text from connection object
        string? connectionString = ConnectionHelpers.GetConnectionString(conn);
        string? commandText = ConnectionHelpers.GetCommandText(conn);

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            throw new InvalidOperationException("Connection has no connection string");
        }

        // Command text can be empty for some connection types (Text, Web)
        if (string.IsNullOrWhiteSpace(commandText))
        {
            commandText = "";
        }

        // Use unified QueryTable creation method
        var createOptions = new PowerQueryHelpers.QueryTableCreateOptions
        {
            Name = options.Name,
            Range = options.Range,
            ConnectionString = connectionString,
            CommandText = commandText,
            ClearWorksheet = options.ClearWorksheet,
            BackgroundQuery = options.BackgroundQuery,
            RefreshOnFileOpen = options.RefreshOnFileOpen,
            SavePassword = options.SavePassword,
            PreserveColumnInfo = options.PreserveColumnInfo,
            PreserveFormatting = options.PreserveFormatting,
            AdjustColumnWidth = options.AdjustColumnWidth,
            RefreshImmediately = options.RefreshImmediately
        };

        PowerQueryHelpers.CreateQueryTable(targetSheet, createOptions);
    }

    #endregion
}

/// <summary>
/// Connection definition for JSON import/export
/// </summary>
internal sealed class ConnectionDefinition
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
