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
    /// Removes QueryTables associated with a query name from a specific worksheet
    /// </summary>
    /// <param name="worksheet">Excel worksheet COM object</param>
    /// <param name="name">Name of the query (spaces will be replaced with underscores for QueryTable names)</param>
    public static void RemoveQueryTablesFromSheet(dynamic worksheet, string name)
    {
        dynamic? queryTables = null;

        try
        {
            queryTables = worksheet.QueryTables;
            string normalizedName = name.Replace(" ", "_");

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
        catch
        {
            // Ignore errors when removing QueryTables - they may not exist
        }
        finally
        {
            ComUtilities.Release(ref queryTables);
        }
    }

    /// <summary>
    /// Unified options for creating QueryTables from any source (Power Query, connections, custom)
    /// </summary>
    public class QueryTableCreateOptions
    {
        /// <summary>
        /// Name of the QueryTable (required)
        /// </summary>
        public required string Name { get; init; }

        /// <summary>
        /// Target range for QueryTable (default: "A1")
        /// </summary>
        public string Range { get; init; } = "A1";

        /// <summary>
        /// Connection string (optional - will be auto-generated for Power Query if not provided)
        /// </summary>
        public string? ConnectionString { get; init; }

        /// <summary>
        /// Command text/SQL query (optional - will be auto-generated for Power Query if not provided)
        /// </summary>
        public string? CommandText { get; init; }

        /// <summary>
        /// Whether to clear worksheet data before creating QueryTable (default: false)
        /// Set to true for Power Query scenarios to prevent column accumulation
        /// </summary>
        public bool ClearWorksheet { get; init; }

        /// <summary>
        /// Whether to refresh data in background (default: false for synchronous behavior)
        /// </summary>
        public bool BackgroundQuery { get; init; }

        /// <summary>
        /// Whether to refresh data when file opens (default: false)
        /// </summary>
        public bool RefreshOnFileOpen { get; init; }

        /// <summary>
        /// Whether to save password in connection (default: false for security)
        /// </summary>
        public bool SavePassword { get; init; }

        /// <summary>
        /// Whether to preserve column information (default: true)
        /// IMPORTANT: Set to FALSE to allow column structure changes when query is updated
        /// If TRUE, column structure is locked at QueryTable creation time
        /// </summary>
        public bool PreserveColumnInfo { get; init; } = true;

        /// <summary>
        /// Whether to preserve formatting (default: true)
        /// </summary>
        public bool PreserveFormatting { get; init; } = true;

        /// <summary>
        /// Whether to auto-adjust column width (default: true)
        /// </summary>
        public bool AdjustColumnWidth { get; init; } = true;

        /// <summary>
        /// Whether to refresh immediately after creation (default: true for immediate feedback)
        /// </summary>
        public bool RefreshImmediately { get; init; } = true;
    }

    /// <summary>
    /// Unified method to create QueryTable from any source (Power Query, connection, custom)
    /// This is the single source of truth for QueryTable creation logic
    /// </summary>
    /// <param name="targetSheet">Target worksheet COM object</param>
    /// <param name="options">QueryTable configuration options</param>
    /// <param name="queryName">Name of Power Query (optional - used to auto-generate connection string if not provided in options)</param>
    public static void CreateQueryTable(dynamic targetSheet, QueryTableCreateOptions options, string? queryName = null)
    {
        dynamic? usedRange = null;
        dynamic? queryTables = null;
        dynamic? queryTable = null;
        dynamic? range = null;

        try
        {
            // Clear worksheet if requested (Power Query scenarios to prevent column accumulation)
            if (options.ClearWorksheet)
            {
                try
                {
                    usedRange = targetSheet.UsedRange;
                    usedRange.Clear();
                }
                catch
                {
                    // Ignore errors if worksheet is empty
                }
            }

            queryTables = targetSheet.QueryTables;
            range = targetSheet.Range[options.Range];

            // Determine connection string and command text
            string connectionString;
            string commandText;

            if (!string.IsNullOrWhiteSpace(options.ConnectionString))
            {
                // Use provided connection string (for connections or custom scenarios)
                connectionString = options.ConnectionString;
                commandText = options.CommandText ?? "";
            }
            else if (!string.IsNullOrWhiteSpace(queryName))
            {
                // Auto-generate for Power Query
                connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
                commandText = $"SELECT * FROM [{queryName}]";
            }
            else
            {
                throw new ArgumentException("Either ConnectionString or queryName must be provided");
            }

            // Create QueryTable
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
            ComUtilities.Release(ref usedRange);
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref queryTables);
        }
    }
}
