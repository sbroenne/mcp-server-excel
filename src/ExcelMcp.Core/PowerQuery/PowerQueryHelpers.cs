using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.PowerQuery;

/// <summary>
/// Helper methods for Power Query operations
/// </summary>
public static class PowerQueryHelpers
{
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
        public bool BackgroundQuery { get; init; } = false;

        /// <summary>
        /// Whether to refresh data when file opens
        /// </summary>
        public bool RefreshOnFileOpen { get; init; } = false;

        /// <summary>
        /// Whether to save password in connection
        /// </summary>
        public bool SavePassword { get; init; } = false;

        /// <summary>
        /// Whether to preserve column information
        /// </summary>
        public bool PreserveColumnInfo { get; init; } = true;

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
        public bool RefreshImmediately { get; init; } = false;
    }

    /// <summary>
    /// Creates a QueryTable connection that loads data from a Power Query to a worksheet
    /// </summary>
    /// <param name="targetSheet">Target worksheet COM object</param>
    /// <param name="queryName">Name of the Power Query</param>
    /// <param name="options">QueryTable configuration options</param>
    public static void CreateQueryTable(dynamic targetSheet, string queryName, QueryTableOptions? options = null)
    {
        options ??= new QueryTableOptions { Name = queryName };

        dynamic? queryTables = null;
        dynamic? queryTable = null;
        dynamic? range = null;

        try
        {
            queryTables = targetSheet.QueryTables;

            // Connection string for Power Query (uses Microsoft.Mashup.OleDb provider)
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"SELECT * FROM [{queryName}]";

            // Create QueryTable at cell A1
            range = targetSheet.Range["A1"];
            queryTable = queryTables.Add(connectionString, range, commandText);

            // Configure QueryTable properties
            queryTable.Name = options.Name.Replace(" ", "_");
            queryTable.RefreshStyle = 1; // xlInsertDeleteCells
            queryTable.BackgroundQuery = options.BackgroundQuery;
            queryTable.RefreshOnFileOpen = options.RefreshOnFileOpen;
            queryTable.SavePassword = options.SavePassword;
            queryTable.PreserveColumnInfo = options.PreserveColumnInfo;
            queryTable.PreserveFormatting = options.PreserveFormatting;
            queryTable.AdjustColumnWidth = options.AdjustColumnWidth;

            // Refresh immediately if requested
            if (options.RefreshImmediately)
            {
                queryTable.Refresh(false);
            }
        }
        finally
        {
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref queryTables);
        }
    }
}
