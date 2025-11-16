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
}
