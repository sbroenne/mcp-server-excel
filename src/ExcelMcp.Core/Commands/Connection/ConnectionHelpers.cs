using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Connections;

/// <summary>
/// Helper methods for Excel connection operations
/// </summary>
public static class ConnectionHelpers
{
    /// <summary>
    /// Gets all connection names from a workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <returns>List of connection names</returns>
    public static List<string> GetConnectionNames(dynamic workbook)
    {
        var names = new List<string>();
        dynamic connections = null!;

        try
        {
            connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic conn = null!;
                try
                {
                    conn = connections.Item(i);
                    string name = conn.Name?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        names.Add(name);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref conn!);
                }
            }
        }
        catch
        {
            // Return empty list if any error occurs
        }
        finally
        {
            ComUtilities.Release(ref connections!);
        }

        return names;
    }

    /// <summary>
    /// Gets the connection type name from XlConnectionType enum value
    /// Per Microsoft docs: https://learn.microsoft.com/en-us/office/vba/api/excel.xlconnectiontype
    /// </summary>
    /// <param name="connectionType">Connection type numeric value</param>
    /// <returns>Human-readable connection type name</returns>
    public static string GetConnectionTypeName(int connectionType)
    {
        return connectionType switch
        {
            1 => "OLEDB",
            2 => "ODBC",
            3 => "TEXT",      // xlConnectionTypeTEXT (was incorrectly "XML")
            4 => "WEB",       // xlConnectionTypeWEB (was incorrectly "Text")
            5 => "XMLMAP",    // xlConnectionTypeXMLMAP
            6 => "DATAFEED",  // xlConnectionTypeDATAFEED
            7 => "MODEL",     // xlConnectionTypeMODEL
            8 => "WORKSHEET", // xlConnectionTypeWORKSHEET
            9 => "NOSOURCE",  // xlConnectionTypeNOSOURCE
            _ => $"Unknown ({connectionType})"
        };
    }



    /// <summary>
    /// Removes connections associated with a query or connection name
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="name">Name of the query or connection</param>
    public static void RemoveConnections(dynamic workbook, string name)
    {
        dynamic connections = null!;

        try
        {
            connections = workbook.Connections;

            // Iterate backwards to safely delete items
            for (int i = connections.Count; i >= 1; i--)
            {
                dynamic conn = null!;
                try
                {
                    conn = connections.Item(i);
                    string connName = conn.Name?.ToString() ?? "";

                    // Match exact name or "Query - Name" pattern
                    if (connName.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                        connName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase))
                    {
                        conn.Delete();
                    }
                }
                finally
                {
                    ComUtilities.Release(ref conn!);
                }
            }
        }
        catch
        {
            // Ignore errors when removing connections - they may not exist
        }
        finally
        {
            ComUtilities.Release(ref connections!);
        }
    }

    /// <summary>
    /// Gets the connection string from a connection object
    /// </summary>
    /// <param name="conn">Excel connection COM object</param>
    /// <returns>Connection string or null if not available</returns>
    public static string? GetConnectionString(dynamic conn)
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

    /// <summary>
    /// Gets the command text from a connection object
    /// </summary>
    /// <param name="conn">Excel connection COM object</param>
    /// <returns>Command text or null if not available</returns>
    public static string? GetCommandText(dynamic conn)
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

    /// <summary>
    /// Gets the command type from a connection object
    /// </summary>
    /// <param name="conn">Excel connection COM object</param>
    /// <returns>Command type string or null if not available</returns>
    public static string? GetCommandType(dynamic conn)
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
                    _ => "Unknown(" + (cmdType.HasValue ? cmdType.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "null") + ")"
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
                    _ => "Unknown(" + (cmdType.HasValue ? cmdType.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "null") + ")"
                };
            }
        }
        catch
        {
            // Property not available
        }

        return null;
    }
}

