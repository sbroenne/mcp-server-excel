namespace Sbroenne.ExcelMcp.Core.Connections;

/// <summary>
/// Helper methods for Excel connection operations
/// </summary>
public static class ConnectionHelpers
{
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
}
