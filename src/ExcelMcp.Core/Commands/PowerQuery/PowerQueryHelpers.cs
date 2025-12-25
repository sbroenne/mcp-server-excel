using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.PowerQuery;

/// <summary>
/// Helper methods for Power Query operations
/// </summary>
public static class PowerQueryHelpers
{
    /// <summary>
    /// Determines if a Power Query connection is orphaned (no corresponding query exists).
    /// An orphaned connection is one that appears to be a Power Query connection (based on
    /// connection string or naming pattern) but has no matching entry in the Queries collection.
    /// This commonly occurs after query deletions, renames, or copy/paste operations in Excel.
    /// 
    /// A connection is considered orphaned if:
    /// 1. It's a Power Query connection (uses Microsoft.Mashup provider)
    /// 2. AND EITHER:
    ///    a. It doesn't follow the standard "Query - {queryName}" naming pattern (e.g., "Connection", "Connection1")
    ///    b. OR it follows the pattern but the corresponding query no longer exists in Workbook.Queries
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="connection">Connection COM object</param>
    /// <returns>True if connection is a Power Query connection with no corresponding query</returns>
    public static bool IsOrphanedPowerQueryConnection(dynamic workbook, dynamic connection)
    {
        // First check if this is even a Power Query connection
        if (!IsPowerQueryConnection(connection))
        {
            return false;
        }

        string connectionName = connection.Name?.ToString() ?? "";

        // Check if connection follows the standard "Query - {queryName}" naming pattern
        // Only connections with this pattern are considered "proper" Power Query connections
        if (!connectionName.StartsWith("Query - ", StringComparison.OrdinalIgnoreCase))
        {
            // Generic names like "Connection", "Connection1", etc. are ALWAYS orphaned
            // even if their Location= points to an existing query.
            // The proper connection for a query is always named "Query - {queryName}".
            return true;
        }

        // Extract the query name from the "Query - {queryName}" pattern
        string expectedQueryName = connectionName["Query - ".Length..];

        // Handle potential suffixes like "Query - Name - Model" (though rare)
        int dashIndex = expectedQueryName.IndexOf(" -", StringComparison.Ordinal);
        if (dashIndex > 0)
        {
            expectedQueryName = expectedQueryName[..dashIndex];
        }

        // Check if a query with this name exists
        dynamic? query = null;
        try
        {
            query = ComUtilities.FindQuery(workbook, expectedQueryName);
            // If query is null, the connection is orphaned
            return query == null;
        }
        finally
        {
            ComUtilities.Release(ref query);
        }
    }

    /// <summary>
    /// Determines if a connection is a Power Query connection
    /// </summary>
    /// <param name="connection">Connection COM object</param>
    /// <returns>True if connection is a Power Query connection</returns>
    public static bool IsPowerQueryConnection(dynamic connection)
    {
        try
        {
            // Power Query connections use Microsoft.Mashup provider
            // Check OLEDBConnection for Mashup provider
            if (connection.Type == 1) // xlConnectionTypeOLEDB
            {
                string connectionString = connection.OLEDBConnection?.Connection?.ToString() ?? "";
                if (connectionString.Contains("Microsoft.Mashup.OleDb", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            // Also check connection name pattern (Power Query connections are named "Query - Name")
            string name = connection.Name?.ToString() ?? "";
            if (name.StartsWith("Query - ", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        catch
        {
            // If any error occurs, assume not a Power Query connection
        }

        return false;
    }

    /// <summary>
    /// Removes QueryTables associated with a query or connection name from all worksheets
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="name">Name of the query or connection (spaces will be replaced with underscores for QueryTable names)</param>
    public static void RemoveQueryTables(dynamic workbook, string name)
    {
        dynamic? worksheets = null;

        try
        {
            worksheets = workbook.Worksheets;
            string normalizedName = name.Replace(" ", "_");

            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic? worksheet = null;
                dynamic? queryTables = null;

                try
                {
                    worksheet = worksheets.Item(ws);
                    queryTables = worksheet.QueryTables;

                    // Iterate backwards to safely delete items
                    for (int qt = queryTables.Count; qt >= 1; qt--)
                    {
                        dynamic? queryTable = null;
                        try
                        {
                            queryTable = queryTables.Item(qt);
                            string queryTableName = queryTable.Name?.ToString() ?? "";

                            // Match QueryTable names that contain the normalized name
                            if (queryTableName.Contains(normalizedName, StringComparison.OrdinalIgnoreCase))
                            {
                                queryTable.Delete();
                            }
                        }
                        finally
                        {
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
        catch
        {
            // Ignore errors when removing QueryTables - they may not exist
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }
    }

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
