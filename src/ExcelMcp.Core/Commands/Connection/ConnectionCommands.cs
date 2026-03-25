using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection management commands - Core data layer (no console output)
/// Provides CRUD operations for Excel data connections (OLEDB, ODBC, Text, Web, etc.)
/// </summary>
public partial class ConnectionCommands : IConnectionCommands
{
    #region Helper Methods

    private static DateTime? GetRefreshDateSafe(object? refreshDate)
    {
        return refreshDate switch
        {
            null => null,
            DateTime dateTime => dateTime,
            double oaDate => DateTime.FromOADate(oaDate),
            _ => null
        };
    }

    private static dynamic? GetTypedSubConnection(dynamic conn)
    {
        int connType = conn.Type;

        if (connType == 1)
        {
            return conn.OLEDBConnection;
        }

        if (connType == 2)
        {
            return conn.ODBCConnection;
        }

        if (connType is 3 or 4)
        {
            try
            {
                return conn.TextConnection;
            }
            catch (COMException)
            {
                try
                {
                    return conn.WebConnection;
                }
                catch (COMException)
                {
                    return null;
                }
            }
        }

        return null;
    }

    private static bool GetBackgroundQuerySetting(dynamic conn)
    {
        try
        {
            int connType = conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.BackgroundQuery ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.BackgroundQuery ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                dynamic? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    return textConnection?.BackgroundQuery ?? false;
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                dynamic? webConnection = null;
                try
                {
                    webConnection = conn.WebConnection;
                    return webConnection?.BackgroundQuery ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException)
        {
        }

        return false;
    }

    private static bool GetRefreshOnFileOpenSetting(dynamic conn)
    {
        try
        {
            int connType = conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.RefreshOnFileOpen ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.RefreshOnFileOpen ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                dynamic? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    return textConnection?.RefreshOnFileOpen ?? false;
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                dynamic? webConnection = null;
                try
                {
                    webConnection = conn.WebConnection;
                    return webConnection?.RefreshOnFileOpen ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException)
        {
        }

        return false;
    }

    private static bool GetSavePasswordSetting(dynamic conn)
    {
        try
        {
            int connType = conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.SavePassword ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.SavePassword ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                dynamic? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    return textConnection?.SavePassword ?? false;
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                dynamic? webConnection = null;
                try
                {
                    webConnection = conn.WebConnection;
                    return webConnection?.SavePassword ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException)
        {
        }

        return false;
    }

    private static int GetRefreshPeriod(dynamic conn)
    {
        try
        {
            int connType = conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.RefreshPeriod ?? 0;
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.RefreshPeriod ?? 0;
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                dynamic? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    return textConnection?.RefreshPeriod ?? 0;
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                dynamic? webConnection = null;
                try
                {
                    webConnection = conn.WebConnection;
                    return webConnection?.RefreshPeriod ?? 0;
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException)
        {
        }

        return 0;
    }

    private static DateTime? GetLastRefreshDate(dynamic conn)
    {
        try
        {
            int connType = conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return GetRefreshDateSafe(oledbConnection?.RefreshDate);
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return GetRefreshDateSafe(odbcConnection?.RefreshDate);
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }
        }
        catch (COMException)
        {
        }

        return null;
    }

    private static string? GetConnectionString(dynamic conn)
    {
        try
        {
            int connType = conn.Type;
            string? connectionString = null;

            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    connectionString = oledbConnection?.Connection?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }
            else if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    connectionString = odbcConnection?.Connection?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }
            else if (connType is 3 or 4)
            {
                dynamic? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    connectionString = textConnection?.Connection?.ToString();
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    dynamic? webConnection = null;
                    try
                    {
                        webConnection = conn.WebConnection;
                        connectionString = webConnection?.Connection?.ToString();
                    }
                    catch (COMException)
                    {
                    }
                    finally
                    {
                        ComUtilities.Release(ref webConnection);
                    }
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
        try
        {
            int connType = conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.CommandText?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.CommandText?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                dynamic? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    return textConnection?.CommandText?.ToString();
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                dynamic? webConnection = null;
                try
                {
                    webConnection = conn.WebConnection;
                    return webConnection?.CommandText?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException)
        {
        }

        return null;
    }

    private static string? GetCommandType(dynamic conn)
    {
        int commandType;
        try
        {
            int connType = conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    if (oledbConnection == null)
                    {
                        return null;
                    }

                    commandType = Convert.ToInt32(oledbConnection.CommandType, CultureInfo.InvariantCulture);
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }
            else if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    if (odbcConnection == null)
                    {
                        return null;
                    }

                    commandType = Convert.ToInt32(odbcConnection.CommandType, CultureInfo.InvariantCulture);
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }
            else
            {
                return null;
            }
        }
        catch (COMException)
        {
            return null;
        }

        return commandType switch
        {
            1 => "Cube",
            2 => "SQL",
            3 => "Table",
            4 => "Default",
            5 => "List",
            _ => $"Unknown({commandType.ToString(CultureInfo.InvariantCulture)})"
        };
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
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    if (oledbConnection != null)
                    {
                        if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                        {
                            oledbConnection.Connection = definition.ConnectionString;
                        }
                        if (!string.IsNullOrWhiteSpace(definition.CommandText))
                        {
                            oledbConnection.CommandText = definition.CommandText;
                        }
                        if (definition.BackgroundQuery.HasValue)
                        {
                            oledbConnection.BackgroundQuery = definition.BackgroundQuery.Value;
                        }
                        if (definition.RefreshOnFileOpen.HasValue)
                        {
                            oledbConnection.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                        }
                        if (definition.SavePassword.HasValue)
                        {
                            oledbConnection.SavePassword = definition.SavePassword.Value;
                        }
                        if (definition.RefreshPeriod.HasValue)
                        {
                            oledbConnection.RefreshPeriod = definition.RefreshPeriod.Value;
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }
            else if (connType == 2) // ODBC
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    if (odbcConnection != null)
                    {
                        if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                        {
                            odbcConnection.Connection = definition.ConnectionString;
                        }
                        if (!string.IsNullOrWhiteSpace(definition.CommandText))
                        {
                            odbcConnection.CommandText = definition.CommandText;
                        }
                        if (definition.BackgroundQuery.HasValue)
                        {
                            odbcConnection.BackgroundQuery = definition.BackgroundQuery.Value;
                        }
                        if (definition.RefreshOnFileOpen.HasValue)
                        {
                            odbcConnection.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                        }
                        if (definition.SavePassword.HasValue)
                        {
                            odbcConnection.SavePassword = definition.SavePassword.Value;
                        }
                        if (definition.RefreshPeriod.HasValue)
                        {
                            odbcConnection.RefreshPeriod = definition.RefreshPeriod.Value;
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }
            else if (connType is 3 or 4) // TEXT (type 3) or WEB (type 4) - Excel may report CSV files as either
            {
                // Excel has type 3/4 confusion: CSV files created with "TEXT;filepath" may be reported as type 4 (WEB)
                // Try TextConnection first (correct for type 3), fall back to WebConnection if that fails
                dynamic? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection; // Try TEXT first
                }
                catch (COMException)
                {
                }

                if (textConnection != null)
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                        {
                            textConnection.Connection = definition.ConnectionString;
                        }
                        if (!string.IsNullOrWhiteSpace(definition.CommandText))
                        {
                            textConnection.CommandText = definition.CommandText;
                        }
                        if (definition.BackgroundQuery.HasValue)
                        {
                            textConnection.BackgroundQuery = definition.BackgroundQuery.Value;
                        }
                        if (definition.RefreshOnFileOpen.HasValue)
                        {
                            textConnection.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                        }
                        if (definition.SavePassword.HasValue)
                        {
                            textConnection.SavePassword = definition.SavePassword.Value;
                        }
                        if (definition.RefreshPeriod.HasValue)
                        {
                            textConnection.RefreshPeriod = definition.RefreshPeriod.Value;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref textConnection);
                    }

                    return;
                }

                dynamic? webConnection = null;
                try
                {
                    webConnection = conn.WebConnection; // Fall back to WEB
                    if (webConnection != null)
                    {
                        if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                        {
                            webConnection.Connection = definition.ConnectionString;
                        }
                        if (!string.IsNullOrWhiteSpace(definition.CommandText))
                        {
                            webConnection.CommandText = definition.CommandText;
                        }
                        if (definition.BackgroundQuery.HasValue)
                        {
                            webConnection.BackgroundQuery = definition.BackgroundQuery.Value;
                        }
                        if (definition.RefreshOnFileOpen.HasValue)
                        {
                            webConnection.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                        }
                        if (definition.SavePassword.HasValue)
                        {
                            webConnection.SavePassword = definition.SavePassword.Value;
                        }
                        if (definition.RefreshPeriod.HasValue)
                        {
                            webConnection.RefreshPeriod = definition.RefreshPeriod.Value;
                        }
                    }
                }
                catch (COMException)
                {
                    // Neither works - skip property updates
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
            else if (connType == 5) // XMLMAP
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

        dynamic? subConn = null;
        try
        {
            subConn = GetTypedSubConnection(conn);
            if (subConn != null)
            {
                SetProperty(subConn, propertyName, value.Value);
            }
        }
        catch (COMException)
        {
            // Property not available for this connection type
        }
        finally
        {
            ComUtilities.Release(ref subConn);
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
        PowerQueryHelpers.QueryTableOptions options,
        CancellationToken cancellationToken)
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
                // Do NOT use EnterLongOperation here. Synchronous QueryTable refresh can depend on
                // inbound Excel callbacks to complete, and rejecting them can deadlock the load.
                OleMessageFilter.SetPendingCancellationToken(cancellationToken);
                try
                {
                    queryTable.Refresh(false);
                }
                finally
                {
                    OleMessageFilter.ClearPendingCancellationToken();
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


