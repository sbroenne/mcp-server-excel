using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.PowerQuery;

/// <summary>
/// Helper methods for Power Query operations
/// </summary>
public static class PowerQueryHelpers
{
    private const string MashupProvider = "Microsoft.Mashup.OleDb.1";

    /// <summary>
    /// Determines if a Power Query connection is orphaned (no corresponding query exists).
    /// Ownership comes from the exact mashup Location property, not the display name.
    /// This commonly occurs after query deletions, renames, or copy/paste operations in Excel.
    /// 
    /// A mashup connection is orphaned when Location is absent or no exact
    /// case-insensitive query name matches that Location.
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="connection">Connection COM object</param>
    /// <returns>True if connection is a Power Query connection with no corresponding query</returns>
    public static bool IsOrphanedPowerQueryConnection(
        Excel.Workbook workbook,
        Excel.WorkbookConnection connection)
    {
        if (!TryGetMashupIdentity(connection, out var identity))
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(identity.Location))
        {
            return true;
        }

        Excel.WorkbookQuery? query = null;
        try
        {
            query = FindQueryByExactName(workbook, identity.Location);
            return query == null;
        }
        finally
        {
            ComUtilities.Release(ref query);
        }
    }

    /// <summary>
    /// Compatibility overload for callers compiled against the object-based API.
    /// </summary>
    public static bool IsOrphanedPowerQueryConnection(object workbook, object connection)
    {
        return IsOrphanedPowerQueryConnection(
            (Excel.Workbook)workbook,
            (Excel.WorkbookConnection)connection);
    }

    /// <summary>
    /// Determines if a connection is a Power Query connection
    /// </summary>
    /// <param name="connection">Connection COM object</param>
    /// <returns>True if connection is a Power Query connection</returns>
    public static bool IsPowerQueryConnection(Excel.WorkbookConnection connection)
    {
        return TryGetMashupIdentity(connection, out _);
    }

    /// <summary>
    /// Compatibility overload for callers compiled against the object-based API.
    /// </summary>
    public static bool IsPowerQueryConnection(object connection)
    {
        return IsPowerQueryConnection((Excel.WorkbookConnection)connection);
    }

    /// <summary>
    /// Determines whether a connection string belongs to the exact mashup query location.
    /// </summary>
    public static bool MatchesMashupLocation(string? connectionString, string queryName)
    {
        return TryParseMashupIdentity(connectionString, out var identity) &&
            string.Equals(identity.Location, queryName, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Reads the exact mashup query location from a workbook connection.
    /// </summary>
    public static bool TryGetMashupLocation(
        Excel.WorkbookConnection connection,
        out string location)
    {
        if (TryGetMashupIdentity(connection, out var identity) &&
            !string.IsNullOrWhiteSpace(identity.Location))
        {
            location = identity.Location;
            return true;
        }

        location = string.Empty;
        return false;
    }

    /// <summary>
    /// Reads mashup identity for read-contract inspection without converting an
    /// unexpected OLEDB COM failure into a false "not loaded" result.
    /// </summary>
    internal static bool TryGetMashupLocationForInspection(
        Excel.WorkbookConnection connection,
        out string location)
    {
        if (Convert.ToInt32(connection.Type, CultureInfo.InvariantCulture) !=
            (int)Excel.XlConnectionType.xlConnectionTypeOLEDB)
        {
            location = string.Empty;
            return false;
        }

        Excel.OLEDBConnection? oleDbConnection = null;
        try
        {
            oleDbConnection = connection.OLEDBConnection;
            string? connectionString = oleDbConnection == null
                ? null
                : Convert.ToString(
                    oleDbConnection.Connection,
                    CultureInfo.InvariantCulture);
            if (oleDbConnection != null &&
                TryParseMashupIdentity(
                    connectionString,
                    out var identity) &&
                !string.IsNullOrWhiteSpace(identity.Location))
            {
                location = identity.Location;
                return true;
            }

            location = string.Empty;
            return false;
        }
        finally
        {
            ComUtilities.Release(ref oleDbConnection);
        }
    }

    /// <summary>
    /// Replaces the exact Location property of a mashup connection string.
    /// </summary>
    public static bool TryReplaceMashupLocation(
        string connectionString,
        string expectedLocation,
        string newLocation,
        out string updatedConnectionString)
    {
        var parsed = ParseConnectionProperties(connectionString);
        if (!parsed.TryGetValue("Provider", out var provider) ||
            !string.Equals(provider.Value, MashupProvider, StringComparison.OrdinalIgnoreCase) ||
            !parsed.TryGetValue("Location", out var location) ||
            !string.Equals(location.Value, expectedLocation, StringComparison.OrdinalIgnoreCase))
        {
            updatedConnectionString = connectionString;
            return false;
        }

        var encodedLocation = EncodePropertyValue(newLocation);
        updatedConnectionString =
            connectionString[..location.ValueStart] +
            encodedLocation +
            connectionString[location.ValueEnd..];
        return true;
    }

    /// <summary>
    /// Removes QueryTables owned by the exact workbook connection.
    /// </summary>
    public static void RemoveQueryTables(
        Excel.Workbook workbook,
        Excel.WorkbookConnection expectedConnection)
    {
        Excel.Sheets? worksheets = null;
        var expectedConnectionName = expectedConnection.Name;
        try
        {
            worksheets = workbook.Worksheets;
            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                Excel.Worksheet? worksheet = null;
                Excel.QueryTables? queryTables = null;
                try
                {
                    worksheet = (Excel.Worksheet)worksheets.Item[ws];
                    queryTables = worksheet.QueryTables;
                    for (int qt = queryTables.Count; qt >= 1; qt--)
                    {
                        Excel.QueryTable? queryTable = null;
                        Excel.WorkbookConnection? queryTableConnection = null;
                        try
                        {
                            queryTable = queryTables.Item(qt);
                            queryTableConnection = queryTable.WorkbookConnection;
                            if (string.Equals(
                                queryTableConnection.Name,
                                expectedConnectionName,
                                StringComparison.OrdinalIgnoreCase))
                            {
                                queryTable.Delete();
                            }
                        }
                        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
                        {
                            // The QueryTable has no accessible WorkbookConnection.
                        }
                        finally
                        {
                            ComUtilities.Release(ref queryTableConnection);
                            ComUtilities.Release(ref queryTable);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref queryTables);
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }
    }

    /// <summary>
    /// Compatibility overload that resolves an exact workbook connection name.
    /// </summary>
    public static void RemoveQueryTables(object workbook, string connectionName)
    {
        var typedWorkbook = (Excel.Workbook)workbook;
        Excel.WorkbookConnection? connection = null;
        try
        {
            connection = FindConnectionByExactName(typedWorkbook, connectionName);
            if (connection != null)
            {
                RemoveQueryTables(typedWorkbook, connection);
            }
        }
        finally
        {
            ComUtilities.Release(ref connection);
        }
    }

    /// <summary>
    /// Removes workbook connections whose mashup Location exactly matches the query.
    /// </summary>
    public static void RemoveConnectionsByMashupLocation(
        Excel.Workbook workbook,
        string queryName)
    {
        Excel.Connections? connections = null;
        var connectionNames = new List<string>();
        try
        {
            connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                Excel.WorkbookConnection? connection = null;
                try
                {
                    connection = connections.Item(i);
                    if (TryGetMashupLocation(connection, out var location) &&
                        string.Equals(location, queryName, StringComparison.OrdinalIgnoreCase))
                    {
                        connectionNames.Add(connection.Name);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connection);
                }
            }

            foreach (var connectionName in connectionNames)
            {
                Excel.WorkbookConnection? connection = null;
                try
                {
                    connection = FindConnectionByExactName(workbook, connectionName);
                    connection?.Delete();
                }
                finally
                {
                    ComUtilities.Release(ref connection);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref connections);
        }
    }

    internal static Excel.WorkbookConnection? FindConnectionByExactName(
        Excel.Workbook workbook,
        string connectionName)
    {
        Excel.Connections? connections = null;
        try
        {
            connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                Excel.WorkbookConnection? connection = null;
                try
                {
                    connection = connections.Item(i);
                    if (string.Equals(
                        connection.Name,
                        connectionName,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        var result = connection;
                        connection = null;
                        return result;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connection);
                }
            }

            return null;
        }
        finally
        {
            ComUtilities.Release(ref connections);
        }
    }

    internal static Excel.WorkbookQuery? FindQueryByExactName(
        Excel.Workbook workbook,
        string queryName)
    {
        Excel.Queries? queries = null;
        try
        {
            queries = workbook.Queries;
            for (int i = 1; i <= queries.Count; i++)
            {
                Excel.WorkbookQuery? query = null;
                try
                {
                    query = queries.Item(i);
                    if (string.Equals(
                        query.Name,
                        queryName,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        var result = query;
                        query = null;
                        return result;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }

            return null;
        }
        finally
        {
            ComUtilities.Release(ref queries);
        }
    }

    private static bool TryGetMashupIdentity(
        Excel.WorkbookConnection connection,
        out MashupIdentity identity)
    {
        Excel.OLEDBConnection? oleDbConnection = null;
        try
        {
            oleDbConnection = connection.OLEDBConnection;
            return TryParseMashupIdentity(
                Convert.ToString(oleDbConnection.Connection, CultureInfo.InvariantCulture),
                out identity);
        }
        catch (COMException)
        {
            identity = default;
            return false;
        }
        finally
        {
            ComUtilities.Release(ref oleDbConnection);
        }
    }

    private static bool TryParseMashupIdentity(
        string? connectionString,
        out MashupIdentity identity)
    {
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            identity = default;
            return false;
        }

        var properties = ParseConnectionProperties(connectionString);
        if (!properties.TryGetValue("Provider", out var provider) ||
            !string.Equals(provider.Value, MashupProvider, StringComparison.OrdinalIgnoreCase))
        {
            identity = default;
            return false;
        }

        properties.TryGetValue("Location", out var location);
        identity = new MashupIdentity(location.Value ?? string.Empty);
        return true;
    }

    private static Dictionary<string, ConnectionProperty> ParseConnectionProperties(
        string connectionString)
    {
        var properties = new Dictionary<string, ConnectionProperty>(
            StringComparer.OrdinalIgnoreCase);
        var segmentStart = 0;
        var quote = '\0';
        var seenEquals = false;
        var valueStarted = false;

        for (int index = 0; index <= connectionString.Length; index++)
        {
            var atEnd = index == connectionString.Length;
            var current = atEnd ? ';' : connectionString[index];

            if (!atEnd && quote != '\0')
            {
                if (current == quote)
                {
                    if (index + 1 < connectionString.Length &&
                        connectionString[index + 1] == quote)
                    {
                        index++;
                    }
                    else
                    {
                        quote = '\0';
                    }
                }

                continue;
            }

            if (!atEnd && current == '=' && !seenEquals)
            {
                seenEquals = true;
                continue;
            }

            if (!atEnd && seenEquals && !valueStarted)
            {
                if (char.IsWhiteSpace(current))
                {
                    continue;
                }

                valueStarted = true;
                if (current is '"' or '\'')
                {
                    quote = current;
                    continue;
                }
            }

            if (current != ';')
            {
                continue;
            }

            AddConnectionProperty(connectionString, segmentStart, index, properties);
            segmentStart = index + 1;
            seenEquals = false;
            valueStarted = false;
        }

        return properties;
    }

    private static void AddConnectionProperty(
        string connectionString,
        int segmentStart,
        int segmentEnd,
        IDictionary<string, ConnectionProperty> properties)
    {
        var equalsIndex = connectionString.IndexOf('=', segmentStart, segmentEnd - segmentStart);
        if (equalsIndex < 0)
        {
            return;
        }

        var key = connectionString[segmentStart..equalsIndex].Trim();
        if (key.Length == 0)
        {
            return;
        }

        var valueStart = equalsIndex + 1;
        while (valueStart < segmentEnd && char.IsWhiteSpace(connectionString[valueStart]))
        {
            valueStart++;
        }

        var valueEnd = segmentEnd;
        while (valueEnd > valueStart &&
            (char.IsWhiteSpace(connectionString[valueEnd - 1]) ||
             connectionString[valueEnd - 1] == '\0'))
        {
            valueEnd--;
        }

        var rawValue = connectionString[valueStart..valueEnd];
        var value = DecodePropertyValue(rawValue);
        properties[key] = new ConnectionProperty(value, valueStart, valueEnd);
    }

    private static string DecodePropertyValue(string rawValue)
    {
        if (rawValue.Length >= 2 &&
            rawValue[0] == rawValue[^1] &&
            rawValue[0] is '"' or '\'')
        {
            var quote = rawValue[0];
            return rawValue[1..^1].Replace(
                new string(quote, 2),
                quote.ToString(),
                StringComparison.Ordinal);
        }

        return rawValue;
    }

    private static string EncodePropertyValue(string value)
    {
        if (value.IndexOfAny([';', '"', '\'']) < 0 &&
            value.Length == value.Trim().Length)
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";
    }

    private readonly record struct MashupIdentity(string Location);

    private readonly record struct ConnectionProperty(
        string? Value,
        int ValueStart,
        int ValueEnd);

    /// <summary>
    /// Options for creating QueryTable connections
    /// </summary>
    public class QueryTableOptions
    {
        /// <summary>
        /// Name of the query or connection
        /// </summary>
        public required string Name { get; init; }

        /// <summary>
        /// Whether to refresh data in background
        /// </summary>
        public bool BackgroundQuery { get; init; }

        /// <summary>
        /// Whether to refresh data when file opens
        /// </summary>
        public bool RefreshOnFileOpen { get; init; }

        /// <summary>
        /// Whether to save password in connection
        /// </summary>
        public bool SavePassword { get; init; }

        /// <summary>
        /// Whether to preserve column information
        /// IMPORTANT: Set to FALSE to allow column structure changes when query is updated
        /// If TRUE, column structure is locked at QueryTable creation time
        /// </summary>
        public bool PreserveColumnInfo { get; init; }

        /// <summary>
        /// Whether to preserve formatting
        /// </summary>
        public bool PreserveFormatting { get; init; } = true;

        /// <summary>
        /// Whether to auto-adjust column width
        /// </summary>
        public bool AdjustColumnWidth { get; init; } = true;

        /// <summary>
        /// Whether to refresh immediately after creation
        /// </summary>
        public bool RefreshImmediately { get; init; }
    }
}
