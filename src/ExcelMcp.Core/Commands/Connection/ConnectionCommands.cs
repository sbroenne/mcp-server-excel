using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.PowerQuery;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection management commands - Core data layer (no console output)
/// Provides CRUD operations for Excel data connections (OLEDB, ODBC, Text, Web, etc.)
/// </summary>
public partial class ConnectionCommands : IConnectionCommands
{
    #region Helper Methods

    /// <summary>
    /// Returns the typed sub-connection (OLEDBConnection, ODBCConnection, TextConnection, or WebConnection)
    /// based on the connection type. For types 3/4, tries TextConnection first then WebConnection
    /// because Excel may report CSV files as either type.
    /// </summary>
    private static dynamic? GetTypedSubConnection(dynamic conn)
    {
        int connType = conn.Type;

        if (connType == 1) return conn.OLEDBConnection;
        if (connType == 2) return conn.ODBCConnection;
        if (connType is 3 or 4)
        {
            try { return conn.TextConnection; }
            catch (COMException)
            {
                try { return conn.WebConnection; }
                catch (COMException) { return null; }
            }
        }

        return null;
    }

    private static bool GetBackgroundQuerySetting(dynamic conn)
    {
        try { return GetTypedSubConnection(conn)?.BackgroundQuery ?? false; }
        catch (COMException) { return false; }
    }

    private static bool GetRefreshOnFileOpenSetting(dynamic conn)
    {
        try { return GetTypedSubConnection(conn)?.RefreshOnFileOpen ?? false; }
        catch (COMException) { return false; }
    }

    private static bool GetSavePasswordSetting(dynamic conn)
    {
        try { return GetTypedSubConnection(conn)?.SavePassword ?? false; }
        catch (COMException) { return false; }
    }

    private static int GetRefreshPeriod(dynamic conn)
    {
        try { return GetTypedSubConnection(conn)?.RefreshPeriod ?? 0; }
        catch (COMException) { return 0; }
    }

    private static DateTime? GetLastRefreshDate(dynamic conn)
    {
        try
        {
            // RefreshDate is only available on OLEDB and ODBC connections
            int connType = conn.Type;
            if (connType is not (1 or 2)) return null;
            return GetTypedSubConnection(conn)?.RefreshDate;
        }
        catch (COMException) { return null; }
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
                dynamic textConn = conn.TextConnection;
                if (textConn != null)
                {
                    connectionString = textConn.Connection?.ToString();
                }
            }
            else if (connType == 5) // WEB (xlConnectionTypeWEB)
            {
                dynamic webConn = conn.WebConnection;
                if (webConn != null)
                {
                    connectionString = webConn.Connection?.ToString();
                }
            }

            // Fallback to root ConnectionString property
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                try
                {
                    connectionString = conn.ConnectionString?.ToString();
                }
                catch (COMException)
                {
                    // Property not available
                }
            }

            return connectionString;
        }
        catch (COMException)
        {
            // Property not available
        }

        return null;
    }

    private static string? GetCommandText(dynamic conn)
    {
        try { return GetTypedSubConnection(conn)?.CommandText?.ToString(); }
        catch (COMException) { return null; }
    }

    private static string? GetCommandType(dynamic conn)
    {
        try
        {
            // CommandType is only available on OLEDB and ODBC connections
            int connType = conn.Type;
            if (connType is not (1 or 2)) return null;

            int? cmdType = GetTypedSubConnection(conn)?.CommandType;
            if (!cmdType.HasValue) return "Unknown(null)";
            return cmdType.Value switch
            {
                1 => "Cube",
                2 => "SQL",
                3 => "Table",
                4 => "Default",
                5 => "List",
                _ => $"Unknown({cmdType.Value.ToString(CultureInfo.InvariantCulture)})"
            };
        }
        catch (COMException) { return null; }
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

        // Reject TEXT/WEB connection strings (legacy, use Power Query or ODC import instead)
        string connStr = definition.ConnectionString.Trim();
        if (connStr.StartsWith("TEXT;", StringComparison.OrdinalIgnoreCase) ||
            connStr.StartsWith("URL;", StringComparison.OrdinalIgnoreCase))
        {
            throw new NotSupportedException(
                "TEXT and WEB connections are no longer supported via create action. " +
                "Use powerquery tool for file/web imports, or create an ODC file and use import-odc action.");
        }

        dynamic? connections = null;
        dynamic? newConnection = null;

        try
        {
            connections = workbook.Connections;

            object commandTypeArgument = DetermineCommandType(definition);

            // Use Add2() method (Add() is deprecated per Microsoft docs)
            // https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.connections.add2
            newConnection = connections.Add2(
                Name: connectionName,
                Description: definition.Description ?? "",
                ConnectionString: definition.ConnectionString,
                CommandText: definition.CommandText ?? "",
                lCmdtype: commandTypeArgument,
                CreateModelConnection: false,         // Don't create PowerPivot model connection
                ImportRelationships: false            // Don't import relationships
            );

            // Connection created successfully - let exceptions propagate naturally
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
                catch (System.Runtime.InteropServices.COMException)
                {
                    try
                    {
                        textOrWeb = conn.WebConnection; // Fall back to WEB
                    }
                    catch (System.Runtime.InteropServices.COMException)
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

    private static object DetermineCommandType(ConnectionDefinition definition)
    {
        if (!string.IsNullOrWhiteSpace(definition.CommandType))
        {
            var value = definition.CommandType.Trim();
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numeric))
            {
                return numeric;
            }

            return value.ToLowerInvariant() switch
            {
                "cube" => 1,
                "sql" => 2,
                "table" => 3,
                "default" => 4,
                "list" => 5,
                _ => Type.Missing
            };
        }

        if (!string.IsNullOrWhiteSpace(definition.CommandText))
        {
            // When command text is provided we default to SQL command type (2).
            return 2;
        }

        return Type.Missing;
    }

    private static void SetConnectionProperty<T>(dynamic conn, string propertyName, T? value) where T : struct
    {
        if (!value.HasValue) return;

        try
        {
            dynamic? subConn = GetTypedSubConnection(conn);
            if (subConn != null)
            {
                SetProperty(subConn, propertyName, value.Value);
            }
        }
        catch (COMException)
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
        catch (System.Runtime.InteropServices.COMException)
        {
            // Property doesn't exist or can't be set
        }
    }

    private static void CreateQueryTableForConnection(
        dynamic targetSheet,
        dynamic conn,
        PowerQueryHelpers.QueryTableOptions options)
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

        dynamic? queryTables = null;
        dynamic? queryTable = null;
        dynamic? range = null;

        try
        {
            queryTables = targetSheet.QueryTables;
            range = targetSheet.Range["A1"];
            queryTable = queryTables.Add(connectionString, range, commandText);

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
                OleMessageFilter.EnterLongOperation();
                try
                {
                    queryTable.Refresh(false);
                }
                finally
                {
                    OleMessageFilter.ExitLongOperation();
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref queryTables);
        }
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


