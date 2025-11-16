using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Connections;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query helper methods (internal utilities)
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Helper method to remove existing query connections and QueryTables
    /// </summary>
    private static void RemoveQueryConnections(dynamic workbook, string queryName)
    {
        dynamic? connections = null;
        dynamic? worksheets = null;
        try
        {
            // Remove connections
            connections = workbook.Connections;
            for (int i = connections.Count; i >= 1; i--)
            {
                dynamic? conn = null;
                try
                {
                    conn = connections.Item(i);
                    string connName = conn.Name?.ToString() ?? "";
                    if (connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                        connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase))
                    {
                        conn.Delete();
                    }
                }
                finally
                {
                    ComUtilities.Release(ref conn);
                }
            }

            // Remove QueryTables
            worksheets = workbook.Worksheets;
            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic? worksheet = null;
                dynamic? queryTables = null;
                try
                {
                    worksheet = worksheets.Item(ws);
                    queryTables = worksheet.QueryTables;

                    for (int qt = queryTables.Count; qt >= 1; qt--)
                    {
                        dynamic? queryTable = null;
                        try
                        {
                            queryTable = queryTables.Item(qt);
                            if (queryTable.Name?.ToString()?.Contains(queryName.Replace(" ", "_")) == true)
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
            // Ignore errors when removing connections
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
            ComUtilities.Release(ref connections);
        }
    }

    /// <summary>
    /// Helper method to create a QueryTable connection that loads data to worksheet
    /// </summary>
    private static void CreateQueryTableConnection(dynamic workbook, dynamic targetSheet, string queryName)
    {
        try
        {
            // Ensure the query exists and is accessible
            dynamic query = ComUtilities.FindQuery(workbook, queryName);
            if (query == null)
            {
                throw new InvalidOperationException($"Query '{queryName}' not found");
            }

            // Get the QueryTables collection
            dynamic queryTables = targetSheet.QueryTables;

            // Build connection string for Power Query
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"SELECT * FROM [{queryName}]";

            // Get the target range - ensure it's valid
            dynamic startRange = targetSheet.Range["A1"];

            // Create the QueryTable
            dynamic queryTable = queryTables.Add(connectionString, startRange, commandText);
            queryTable.Name = queryName.Replace(" ", "_");
            queryTable.RefreshStyle = 1; // xlInsertDeleteCells
            queryTable.BackgroundQuery = false;
            queryTable.PreserveColumnInfo = false;  // Allow column structure changes when M code updates
            queryTable.PreserveFormatting = true;
            queryTable.AdjustColumnWidth = true;
            queryTable.RefreshOnFileOpen = false;
            queryTable.SavePassword = false;

            // Refresh to load data immediately
            queryTable.Refresh(false);
        }
        catch (COMException comEx)
        {
            // Provide more detailed error information
            string hexCode = $"0x{comEx.HResult:X}";
            throw new InvalidOperationException(
                $"Failed to create QueryTable connection for '{queryName}': {comEx.Message} (Error: {hexCode}). " +
                $"This may occur if the query needs to be refreshed first or if there are data source connectivity issues.",
                comEx);
        }
    }

    /// <summary>
    /// Configures a Power Query to load to Data Model using Excel COM API
    /// Based on validated VBA pattern using Connections.Add2 method
    /// Reference: Working VBA code that successfully loads queries to Data Model
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="queryName">Name of the query to configure</param>
    /// <param name="errorMessage">Detailed error message if configuration fails</param>
    /// <returns>True if configuration succeeded, false if exception caught</returns>
    private static bool SetQueryLoadToDataModel(dynamic workbook, string queryName, out string? errorMessage)
    {
        dynamic? connections = null;
        dynamic? newConnection = null;
        errorMessage = null;

        try
        {
            connections = workbook.Connections;

            // Remove existing connections for this query to avoid conflicts
            ConnectionHelpers.RemoveConnections(workbook, queryName);

            // Use Connections.Add2 method (Excel 2013+) with Data Model parameters
            // This is the Microsoft-documented approach for loading Power Query to Data Model
            // Based on working VBA pattern:
            // w.Connections.Add2 "Query - " & query.Name, _
            //     "Connection to the '" & query.Name & "' query in the workbook.", _
            //     "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.Name, _
            //     """" & query.Name & """", 6, True, False

            string connectionName = $"Query - {queryName}";
            string description = $"Connection to the '{queryName}' query in the workbook.";
            string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
            string commandText = $"\"{queryName}\""; // Quoted query name
            int commandType = 6; // Data Model command type (xlCmdDAX or similar)
            bool createModelConnection = true; // CRITICAL: This loads to Data Model
            bool importRelationships = false;

            newConnection = connections.Add2(
                connectionName,
                description,
                connectionString,
                commandText,
                commandType,
                createModelConnection,
                importRelationships
            );

            return true;
        }
        catch (Exception ex)
        {
            // Capture detailed error for user feedback
            errorMessage = ex.Message;
            System.Diagnostics.Debug.WriteLine($"Failed to configure Data Model loading: {ex.Message}");
            return false;
        }
        finally
        {
            ComUtilities.Release(ref newConnection);
            ComUtilities.Release(ref connections);
        }
    }

    /// <summary>
    /// Check if a query is configured for data model loading
    /// </summary>
    private static bool CheckQueryDataModelConfiguration(dynamic query)
    {
        try
        {
            // Method 1: Check if the query has LoadToWorksheetModel property set
            try
            {
                bool loadToModel = query.LoadToWorksheetModel;
                if (loadToModel) return true;
            }
            catch
            {
                // Property doesn't exist
            }

            // Method 2: Check if query has ModelConnection property
            try
            {
                dynamic modelConnection = query.ModelConnection;
                if (modelConnection != null) return true;
            }
            catch
            {
                // Property doesn't exist
            }

            // Since we now use explicit DataModel_ connection markers,
            // this method is mainly for detecting native Excel data model configurations
            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Core connection refresh logic - finds and refreshes the connection for a query
    /// Can be called from both UpdateAsync and RefreshAsync to avoid code duplication
    /// </summary>
    private static void RefreshConnectionByQueryName(dynamic workbook, string queryName)
    {
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
                targetConnection.Refresh();
            }
            finally
            {
                ComUtilities.Release(ref targetConnection);
            }
        }
    }
}
