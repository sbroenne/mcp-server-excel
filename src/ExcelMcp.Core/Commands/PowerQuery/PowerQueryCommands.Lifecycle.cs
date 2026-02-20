using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query lifecycle operations (List, View, Import, Export, Update, Delete)
/// </summary>
public partial class PowerQueryCommands
{
    /// <inheritdoc />
    public PowerQueryListResult List(IExcelBatch batch)
    {
        var result = new PowerQueryListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Queries? queriesCollection = null;
            try
            {
                queriesCollection = ctx.Book.Queries;
                int count = queriesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    Excel.WorkbookQuery? query = null;
                    try
                    {
                        query = queriesCollection.Item(i);
                        string name = query.Name ?? $"Query{i}";

                        // Try to read formula - some queries may not have accessible formulas
                        string formula = "";
                        try
                        {
                            formula = query.Formula?.ToString() ?? "";
                        }
                        catch (COMException)
                        {
                            // Formula property not accessible (e.g., corrupted query, permission issue)
                            // Don't fail the entire List operation - just mark this query
                            formula = "";
                        }

                        string preview = formula.Length > 80 ? formula[..77] + "..." : formula;
                        if (string.IsNullOrEmpty(formula))
                        {
                            preview = "(formula not accessible)";
                        }

                        // Check if loaded to table (ListObject) - same pattern as GetLoadConfig
                        bool isConnectionOnly = true;
                        dynamic? worksheets = null;
                        try
                        {
                            worksheets = ctx.Book.Worksheets;
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
                                        dynamic? wbConn = null;
                                        dynamic? oledbConn = null;
                                        try
                                        {
                                            listObject = listObjects.Item(lo);

                                            // QueryTable property may throw 0x800A03EC if ListObject doesn't have a valid QueryTable
                                            // This is normal - not all ListObjects have QueryTables (e.g., manually created tables)
                                            try
                                            {
                                                queryTable = listObject.QueryTable;
                                            }
                                            catch (COMException ex)
                                                when (ex.HResult == unchecked((int)0x800A03EC))
                                            {
                                                // ListObject doesn't have QueryTable - skip it
                                                continue;
                                            }

                                            if (queryTable == null) continue;

                                            wbConn = queryTable.WorkbookConnection;
                                            if (wbConn == null) continue;

                                            // Non-OLEDB connection types (Type=7, Type=8) throw COMException
                                            try { oledbConn = wbConn.OLEDBConnection; }
                                            catch (System.Runtime.InteropServices.COMException) { continue; }
                                            if (oledbConn == null) continue;

                                            string connString = oledbConn.Connection?.ToString() ?? "";
                                            bool isMashup = connString.Contains("Provider=Microsoft.Mashup.OleDb.1", StringComparison.OrdinalIgnoreCase);
                                            bool locationMatches = connString.Contains($"Location={name}", StringComparison.OrdinalIgnoreCase);

                                            if (isMashup && locationMatches)
                                            {
                                                isConnectionOnly = false;
                                                ComUtilities.Release(ref oledbConn!);
                                                ComUtilities.Release(ref wbConn!);
                                                ComUtilities.Release(ref queryTable!);
                                                ComUtilities.Release(ref listObject!);
                                                break;
                                            }
                                        }
                                        finally
                                        {
                                            ComUtilities.Release(ref oledbConn!);
                                            ComUtilities.Release(ref wbConn!);
                                            ComUtilities.Release(ref queryTable!);
                                            ComUtilities.Release(ref listObject!);
                                        }
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref listObjects!);
                                    ComUtilities.Release(ref worksheet!);
                                }
                                if (!isConnectionOnly) break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref worksheets!);
                        }

                        // Also check for Data Model connections
                        // A query loaded ONLY to Data Model has no ListObjects but has a connection
                        if (isConnectionOnly)
                        {
                            dynamic? connections = null;
                            try
                            {
                                connections = ctx.Book.Connections;
                                for (int c = 1; c <= connections.Count; c++)
                                {
                                    dynamic? conn = null;
                                    try
                                    {
                                        conn = connections.Item(c);
                                        string connName = conn.Name?.ToString() ?? "";

                                        // Check if this is a Data Model connection for our query
                                        // Patterns:
                                        // - "Query - {queryName}" (worksheet connection)
                                        // - "Query - {queryName} (Data Model)" (Data Model connection)
                                        // - "Query - {queryName} - suffix" (legacy pattern)
                                        if (connName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase) ||
                                            connName.StartsWith($"Query - {name} -", StringComparison.OrdinalIgnoreCase) ||
                                            connName.StartsWith($"Query - {name} (", StringComparison.OrdinalIgnoreCase))
                                        {
                                            // Has Data Model connection - NOT connection-only
                                            isConnectionOnly = false;
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
                        }

                        result.Queries.Add(new PowerQueryInfo
                        {
                            Name = name,
                            Formula = formula,
                            FormulaPreview = preview,
                            IsConnectionOnly = isConnectionOnly
                        });
                    }
                    catch (COMException)
                    {
                        // Skip query if COM error occurs during processing
                        // This allows listing to continue for remaining queries
                        // COM exceptions occur for corrupted queries or access issues
                        continue;
                    }
                    finally
                    {
                        ComUtilities.Release(ref query);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref queriesCollection);
            }
        });
    }

    // View method moved to PowerQueryCommands.View.cs (standalone implementation)

    /// <inheritdoc />
    public PowerQueryLoadConfigResult GetLoadConfig(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryLoadConfigResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookQuery? query = null;
            dynamic? worksheets = null;
            dynamic? connections = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // Check for ListObjects first (Power Query loaded to table creates a ListObject)
                bool hasTableConnection = false;
                bool hasDataModelConnection = false;
                string? targetSheet = null;

                worksheets = ctx.Book.Worksheets;
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
                            dynamic? wbConn = null;
                            dynamic? oledbConn = null;
                            try
                            {
                                listObject = listObjects.Item(lo);
                                queryTable = listObject.QueryTable;
                                if (queryTable == null) continue;

                                wbConn = queryTable.WorkbookConnection;
                                if (wbConn == null) continue;

                                // Non-OLEDB connection types (Type=7, Type=8) throw COMException
                                try { oledbConn = wbConn.OLEDBConnection; }
                                catch (System.Runtime.InteropServices.COMException) { continue; }
                                if (oledbConn == null) continue;

                                string connString = oledbConn.Connection?.ToString() ?? "";
                                bool isMashup = connString.Contains("Provider=Microsoft.Mashup.OleDb.1", StringComparison.OrdinalIgnoreCase);
                                bool locationMatches = connString.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase);

                                // Also check CommandText as fallback
                                string commandText = "";
                                try
                                {
                                    if (queryTable.CommandText is object[] arr && arr.Length > 0)
                                        commandText = arr[0]?.ToString() ?? "";
                                    else
                                        commandText = queryTable.CommandText?.ToString() ?? "";
                                }
                                catch (COMException)
                                {
                                    // CommandText property may not be accessible for certain QueryTable types
                                }

                                bool cmdMatches = commandText.Contains($"[{queryName}]", StringComparison.OrdinalIgnoreCase);

                                if (isMashup && (locationMatches || cmdMatches))
                                {
                                    hasTableConnection = true;
                                    targetSheet = worksheet.Name;
                                    ComUtilities.Release(ref oledbConn!);
                                    ComUtilities.Release(ref wbConn!);
                                    ComUtilities.Release(ref queryTable!);
                                    ComUtilities.Release(ref listObject!);
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref oledbConn!);
                                ComUtilities.Release(ref wbConn!);
                                ComUtilities.Release(ref queryTable!);
                                ComUtilities.Release(ref listObject!);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref listObjects!);
                        ComUtilities.Release(ref worksheet!);
                    }
                    if (hasTableConnection) break;
                }

                // Check connections for Data Model membership using InModel property
                connections = ctx.Book.Connections;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic? conn = null;
                    dynamic? oledbConn = null;
                    try
                    {
                        conn = connections.Item(i);
                        string connName = conn.Name?.ToString() ?? "";

                        // Check if this connection is related to our query
                        // Patterns:
                        // - "{queryName}" (exact match)
                        // - "Query - {queryName}" (worksheet connection)
                        // - "Query - {queryName} (Data Model)" (Data Model connection)
                        // - "Query - {queryName} - suffix" (legacy pattern)
                        bool isQueryConnection = connName.Equals(queryName, StringComparison.OrdinalIgnoreCase) ||
                            connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase) ||
                            connName.StartsWith($"Query - {queryName} -", StringComparison.OrdinalIgnoreCase) ||
                            connName.StartsWith($"Query - {queryName} (", StringComparison.OrdinalIgnoreCase);

                        // Also check connection string for Power Query pattern
                        if (!isQueryConnection)
                        {
                            try
                            {
                                oledbConn = conn.OLEDBConnection;
                                if (oledbConn != null)
                                {
                                    string connString = oledbConn.Connection?.ToString() ?? "";
                                    bool isPowerQuery = connString.Contains("Provider=Microsoft.Mashup.OleDb.1", StringComparison.OrdinalIgnoreCase);
                                    bool matchesQuery = connString.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase);
                                    isQueryConnection = isPowerQuery && matchesQuery;
                                }
                            }
                            catch (Exception ex) when (ex is COMException or System.Reflection.TargetInvocationException)
                            {
                                // Connection type doesn't have OLEDBConnection property - skip
                            }
                        }

                        if (isQueryConnection)
                        {
                            result.HasConnection = true;

                            // Check InModel property to detect Data Model connections
                            try
                            {
                                bool inModel = conn.InModel;
                                if (inModel)
                                {
                                    hasDataModelConnection = true;
                                }
                            }
                            catch (Exception ex) when (ex is COMException or System.Reflection.TargetInvocationException)
                            {
                                // InModel property not available for this connection type
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref oledbConn!);
                        ComUtilities.Release(ref conn);
                    }
                }

                // Determine load mode
                if (hasTableConnection && hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToBoth;
                }
                else if (hasTableConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToTable;
                }
                else if (hasDataModelConnection)
                {
                    result.LoadMode = PowerQueryLoadMode.LoadToDataModel;
                }
                else
                {
                    result.LoadMode = PowerQueryLoadMode.ConnectionOnly;
                }

                result.TargetSheet = targetSheet;
                result.IsLoadedToDataModel = hasDataModelConnection;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string queryName)
    {
        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookQuery? query = null;
            Excel.Queries? queriesCollection = null;
            dynamic? worksheets = null;

            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // STEP 1: Clean up any ListObjects (tables) that reference this query
                // When a query is loaded to a worksheet, Excel creates a ListObject with QueryTable
                // Delete must remove these to prevent orphaned tables
                worksheets = ctx.Book.Worksheets;
                int worksheetCount = worksheets.Count;

                for (int i = 1; i <= worksheetCount; i++)
                {
                    dynamic? sheet = null;
                    dynamic? listObjects = null;

                    try
                    {
                        sheet = worksheets.Item(i);
                        listObjects = sheet.ListObjects;
                        int tableCount = listObjects.Count;

                        // Iterate backwards to safely delete while iterating
                        for (int j = tableCount; j >= 1; j--)
                        {
                            dynamic? table = null;
                            dynamic? queryTable = null;
                            dynamic? oleDbConnection = null;

                            try
                            {
                                table = listObjects.Item(j);

                                // Check if this table has a QueryTable with our query
                                try
                                {
                                    queryTable = table.QueryTable;
                                    if (queryTable != null)
                                    {
                                        oleDbConnection = queryTable.WorkbookConnection?.OLEDBConnection;
                                        if (oleDbConnection != null)
                                        {
                                            string? connString = oleDbConnection.Connection?.ToString() ?? "";
                                            // Check if connection string references our query
                                            // Format: "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=QueryName"
                                            if (connString.Contains("Microsoft.Mashup.OleDb") &&
                                                connString.Contains($"Location={queryName}"))
                                            {
                                                // This table is associated with our query - delete it
                                                table.Delete();
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex) when (ex is COMException or System.Reflection.TargetInvocationException)
                                {
                                    // Table doesn't have QueryTable property - skip
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref oleDbConnection);
                                ComUtilities.Release(ref queryTable);
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref listObjects);
                        ComUtilities.Release(ref sheet);
                    }
                }

                // STEP 2: Remove Data Model connections
                // Data Model connections follow pattern: "Query - {queryName}" or "Query - {queryName} - suffix"
                dynamic? connections = null;
                try
                {
                    connections = ctx.Book.Connections;
                    var connectionsToDelete = new List<string>();

                    for (int c = 1; c <= connections.Count; c++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(c);
                            string connName = conn.Name?.ToString() ?? "";

                            // Check if this is a connection for our query
                            // Patterns:
                            // - "Query - {queryName}" (worksheet connection)
                            // - "Query - {queryName} (Data Model)" (Data Model connection)
                            // - "Query - {queryName} - suffix" (legacy pattern)
                            if (connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase) ||
                                connName.StartsWith($"Query - {queryName} -", StringComparison.OrdinalIgnoreCase) ||
                                connName.StartsWith($"Query - {queryName} (", StringComparison.OrdinalIgnoreCase))
                            {
                                connectionsToDelete.Add(connName);
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }

                    // Delete connections
                    foreach (var connName in connectionsToDelete)
                    {
                        dynamic? connToDelete = null;
                        try
                        {
                            connToDelete = connections.Item(connName);
                            connToDelete.Delete();
                        }
                        catch (COMException)
                        {
                            // Connection may have already been deleted - safe to ignore
                        }
                        finally
                        {
                            ComUtilities.Release(ref connToDelete);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                queriesCollection = ctx.Book.Queries;
                queriesCollection.Item(queryName).Delete();

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref queriesCollection);
                ComUtilities.Release(ref query);
            }
        });
    }


    /// <summary>
    /// Converts query to connection-only (removes data load)
    /// Uses ListObjects pattern (matches Delete cleanup logic)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of the query</param>
    /// <returns>Operation result</returns>
    public OperationResult Unload(IExcelBatch batch, string queryName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "unload"
        };

        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookQuery? query = null;
            dynamic? worksheets = null;

            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // Remove ListObjects (tables) that reference this query
                // Same pattern as Delete cleanup
                worksheets = ctx.Book.Worksheets;
                int worksheetCount = worksheets.Count;

                for (int i = 1; i <= worksheetCount; i++)
                {
                    dynamic? sheet = null;
                    dynamic? listObjects = null;

                    try
                    {
                        sheet = worksheets.Item(i);
                        listObjects = sheet.ListObjects;
                        int tableCount = listObjects.Count;

                        // Iterate backwards to safely delete while iterating
                        for (int j = tableCount; j >= 1; j--)
                        {
                            dynamic? table = null;
                            dynamic? queryTable = null;
                            dynamic? oleDbConnection = null;

                            try
                            {
                                table = listObjects.Item(j);

                                // Check if this table has a QueryTable with our query
                                try
                                {
                                    queryTable = table.QueryTable;
                                    if (queryTable != null)
                                    {
                                        oleDbConnection = queryTable.WorkbookConnection?.OLEDBConnection;
                                        if (oleDbConnection != null)
                                        {
                                            string? connString = oleDbConnection.Connection?.ToString() ?? "";
                                            // Check if connection string references our query
                                            if (connString.Contains("Microsoft.Mashup.OleDb") &&
                                                connString.Contains($"Location={queryName}"))
                                            {
                                                // This table is associated with our query - delete it
                                                table.Delete();
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex) when (ex is COMException or System.Reflection.TargetInvocationException)
                                {
                                    // Table doesn't have QueryTable property - skip
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref oleDbConnection);
                                ComUtilities.Release(ref queryTable);
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref listObjects);
                        ComUtilities.Release(ref sheet);
                    }
                }

                // STEP 2: Remove Data Model connections
                // Data Model connections follow pattern: "Query - {queryName}" or "Query - {queryName} - suffix"
                dynamic? connections = null;
                try
                {
                    connections = ctx.Book.Connections;
                    var connectionsToDelete = new List<string>();

                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(i);
                            string connName = conn.Name?.ToString() ?? "";

                            // Check if this is a connection for our query
                            // Patterns:
                            // - "Query - {queryName}" (worksheet connection)
                            // - "Query - {queryName} (Data Model)" (Data Model connection)
                            // - "Query - {queryName} - suffix" (legacy pattern)
                            if (connName.Equals($"Query - {queryName}", StringComparison.OrdinalIgnoreCase) ||
                                connName.StartsWith($"Query - {queryName} -", StringComparison.OrdinalIgnoreCase) ||
                                connName.StartsWith($"Query - {queryName} (", StringComparison.OrdinalIgnoreCase))
                            {
                                connectionsToDelete.Add(connName);
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }

                    // Delete connections (must iterate separately to avoid modifying collection while enumerating)
                    foreach (var connName in connectionsToDelete)
                    {
                        dynamic? connToDelete = null;
                        try
                        {
                            connToDelete = connections.Item(connName);
                            connToDelete.Delete();
                        }
                        catch (COMException)
                        {
                            // Connection may have already been deleted or is in use - safe to ignore
                        }
                        finally
                        {
                            ComUtilities.Release(ref connToDelete);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref connections);
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref query);
            }
        }, cancellationToken: default);
    }
}



