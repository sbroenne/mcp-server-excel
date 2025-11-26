using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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
            dynamic? queriesCollection = null;
            try
            {
                queriesCollection = ctx.Book.Queries;
                int count = queriesCollection.Count;

                for (int i = 1; i <= count; i++)
                {
                    dynamic? query = null;
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
                        catch (System.Runtime.InteropServices.COMException)
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
                                            catch (System.Runtime.InteropServices.COMException ex)
                                                when (ex.HResult == unchecked((int)0x800A03EC))
                                            {
                                                // ListObject doesn't have QueryTable - skip it
                                                continue;
                                            }

                                            if (queryTable == null) continue;

                                            wbConn = queryTable.WorkbookConnection;
                                            if (wbConn == null) continue;

                                            oledbConn = wbConn.OLEDBConnection;
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
            dynamic? query = null;
            dynamic? worksheets = null;
            dynamic? connections = null;
            dynamic? names = null;
            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found");
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

                                oledbConn = wbConn.OLEDBConnection;
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
                                catch (System.Runtime.InteropServices.COMException)
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

                // Check for connections (for data model or other types)
                connections = ctx.Book.Connections;
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
                            result.HasConnection = true;

                            // If we don't have a table connection but have a workbook connection,
                            // it's likely a data model connection
                            if (!hasTableConnection)
                            {
                                hasDataModelConnection = true;
                            }
                        }
                        else if (connName.Equals($"DataModel_{queryName}", StringComparison.OrdinalIgnoreCase))
                        {
                            // This is our explicit data model connection marker
                            result.HasConnection = true;
                            hasDataModelConnection = true;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref conn);
                    }
                }

                // Always check for named range markers that indicate data model loading
                // (even if we have table connections, for LoadToBoth mode)
                if (!hasDataModelConnection)
                {
                    // Check for our data model marker
                    try
                    {
                        names = ctx.Book.Names;
                        string markerName = $"DataModel_Query_{queryName}";

                        for (int i = 1; i <= names.Count; i++)
                        {
                            dynamic? existingName = null;
                            try
                            {
                                existingName = names.Item(i);
                                if (existingName.Name.ToString() == markerName)
                                {
                                    hasDataModelConnection = true;
                                    ComUtilities.Release(ref existingName);
                                    break;
                                }
                            }
                            finally
                            {
                                ComUtilities.Release(ref existingName);
                            }
                        }
                    }
                    catch
                    {
                        // Cannot check names
                    }

                    // Fallback: Check if the query has data model indicators
                    if (!hasDataModelConnection)
                    {
                        hasDataModelConnection = CheckQueryDataModelConfiguration(query);
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
                ComUtilities.Release(ref names);
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref worksheets);
                ComUtilities.Release(ref query);
            }
        });
    }

    /// <inheritdoc />
    public void Delete(IExcelBatch batch, string queryName)
    {
        // Validate query name
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        batch.Execute((ctx, ct) =>
        {
            dynamic? query = null;
            dynamic? queriesCollection = null;
            dynamic? worksheets = null;

            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found");
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
                                catch
                                {
                                    // Table might not have QueryTable - skip
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

                // STEP 2: Delete the query itself
                queriesCollection = ctx.Book.Queries;
                queriesCollection.Item(queryName).Delete();

                return 0;
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
    /// Helper to get all query names
    /// </summary>
    private static List<string> GetQueryNames(dynamic workbook)
    {
        var names = new List<string>();
        dynamic? queriesCollection = null;
        try
        {
            queriesCollection = workbook.Queries;
            for (int i = 1; i <= queriesCollection.Count; i++)
            {
                dynamic? query = null;
                try
                {
                    query = queriesCollection.Item(i);
                    names.Add(query.Name);
                }
                finally
                {
                    ComUtilities.Release(ref query);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref queriesCollection);
        }
        return names;
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
            dynamic? query = null;
            dynamic? worksheets = null;

            try
            {
                query = ComUtilities.FindQuery(ctx.Book, queryName);
                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found");
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
                                catch
                                {
                                    // Table might not have QueryTable - skip
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

