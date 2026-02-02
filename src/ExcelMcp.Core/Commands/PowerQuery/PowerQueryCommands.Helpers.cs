using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query helper methods (internal utilities)
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Core connection refresh logic - finds and refreshes the connection for a query.
    ///
    /// Error propagation depends on connection type:
    /// - Worksheet queries (InModel=false): Errors thrown via QueryTable.Refresh(false)
    /// - Data Model queries (InModel=true): Errors thrown via Connection.Refresh()
    ///
    /// Strategy order ensures we use the appropriate method for each connection type:
    /// 1. Try QueryTable.Refresh() first (handles worksheet queries)
    /// 2. Fall back to Connection.Refresh() (handles Data Model queries)
    /// </summary>
    /// <returns>True if refresh was executed, false if no connection or table found</returns>
    /// <exception cref="Exception">Thrown if Power Query has formula errors</exception>
    private static bool RefreshConnectionByQueryName(dynamic workbook, string queryName)
    {
        // Strategy 1: Find and refresh QueryTable directly on worksheet
        // For worksheet queries (InModel=false), errors are thrown by QueryTable.Refresh()
        if (RefreshQueryTableByName(workbook, queryName))
        {
            return true;
        }

        // Strategy 2: Find connection by name patterns and refresh
        // For Data Model queries (InModel=true), errors are thrown by Connection.Refresh()
        dynamic? targetConnection = null;
        dynamic? connections = null;
        try
        {
            connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic? conn = null;
                try
                {
                    conn = connections.Item(i);
                    string connName = conn.Name?.ToString() ?? "";
                    if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                        connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        targetConnection = conn;
                        conn = null; // Don't release - we're using it
                        break;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref conn);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref connections);
        }

        if (targetConnection != null)
        {
            try
            {
                // For Data Model connections, this throws on Power Query errors
                targetConnection.Refresh();
                return true;
            }
            finally
            {
                ComUtilities.Release(ref targetConnection);
            }
        }

        return false;
    }

    /// <summary>
    /// Finds and refreshes a QueryTable by searching ListObjects on all worksheets.
    /// Matches by query name in the QueryTable's connection string (Location=queryName).
    /// </summary>
    /// <returns>True if QueryTable was found and refreshed</returns>
    /// <exception cref="Exception">Thrown if Power Query has formula errors</exception>
    private static bool RefreshQueryTableByName(dynamic workbook, string queryName)
    {
        dynamic? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;

            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic? worksheet = null;
                dynamic? listObjects = null;
                try
                {
                    worksheet = worksheets.Item(ws);
                    listObjects = worksheet.ListObjects;

                    for (int lo = 1; lo <= listObjects.Count; lo++)
                    {
                        dynamic? listObject = null;
                        dynamic? queryTable = null;
                        try
                        {
                            listObject = listObjects.Item(lo);

                            // Try to get QueryTable - not all ListObjects have one
                            try
                            {
                                queryTable = listObject.QueryTable;
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                                // ListObject doesn't have a QueryTable - expected for user-created tables
                                continue;
                            }

                            if (queryTable == null)
                            {
                                continue;
                            }

                            // Check if this QueryTable is for our query by examining connection string
                            // Format: "OLEDB;...;Location=QueryName;..."
                            string? connection = queryTable.Connection?.ToString();
                            if (connection != null &&
                                connection.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase))
                            {
                                // Found it! Refresh and let any Power Query errors propagate
                                queryTable.Refresh(false); // Synchronous refresh - THROWS on error
                                return true;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref queryTable);
                            ComUtilities.Release(ref listObject);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listObjects);
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }

        return false;
    }
}
