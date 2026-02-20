using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Formatting;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query Update operations - STANDALONE implementation.
/// Does NOT use any existing helper methods.
/// Based on external reference code pattern.
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Update Power Query M code. Preserves load configuration (worksheet/data model).
    /// M code is automatically formatted using the powerqueryformatter.com API before saving.
    /// - Worksheet queries: Uses QueryTable.Refresh(false) for synchronous refresh with column propagation
    /// - Data Model queries: Uses connection.Refresh() to update the Data Model
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of query to update</param>
    /// <param name="mCode">New M code</param>
    /// <param name="refresh">Whether to refresh data after update (default: true)</param>
    /// <exception cref="ArgumentException">Thrown when queryName or mCode is invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when query not found or update fails</exception>
    public OperationResult Update(IExcelBatch batch, string queryName, string mCode, bool refresh = true)
    {
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        if (string.IsNullOrWhiteSpace(mCode))
        {
            throw new ArgumentException("M code cannot be empty", nameof(mCode));
        }

        // Format M code before saving (outside batch.Execute for async operation)
        // Formatting is done synchronously to maintain method signature compatibility
        // Falls back to original if formatting fails
        string formattedMCode = MCodeFormatter.FormatAsync(mCode).GetAwaiter().GetResult();

        return batch.Execute((ctx, ct) =>
        {
            Excel.Queries? queries = null;
            Excel.WorkbookQuery? query = null;
            dynamic? worksheets = null;
            dynamic? targetWorksheet = null;
            dynamic? existingQueryTable = null;

            try
            {
                // STEP 1: Find the Power Query
                queries = ctx.Book.Queries;
                query = null;
                for (int i = 1; i <= queries.Count; i++)
                {
                    dynamic? q = null;
                    try
                    {
                        q = queries.Item(i);
                        string qName = q.Name?.ToString() ?? "";
                        if (qName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                        {
                            query = q;
                            q = null; // Don't release - we're keeping the reference
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref q!);
                    }
                }

                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // STEP 2: Find existing QueryTable (preferred) or ListObject bound to this query
                // Pattern 1: QueryTable created by LoadTo/Create (uses QueryTables.Add)
                // Pattern 2: ListObject created by previous Update (uses ListObjects.Add)
                bool foundQueryTable = false;

                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);

                        // FIRST: Check for QueryTable (Pattern 1 - from LoadTo/Create)
                        dynamic? queryTables = null;
                        try
                        {
                            queryTables = worksheet.QueryTables;
                            for (int qt = 1; qt <= queryTables.Count; qt++)
                            {
                                dynamic? qTable = null;
                                dynamic? wbConn = null;
                                dynamic? oledbConn = null;
                                try
                                {
                                    qTable = queryTables.Item(qt);
                                    wbConn = qTable.WorkbookConnection;
                                    if (wbConn == null) continue;

                                    // NOTE: Accessing OLEDBConnection on non-OLEDB connection types
                                    // (e.g., Type=7 ThisWorkbookDataModel, Type=8 workbook connections)
                                    // throws COMException 0x800A03EC. We must catch and skip.
                                    try
                                    {
                                        oledbConn = wbConn.OLEDBConnection;
                                    }
                                    catch (System.Runtime.InteropServices.COMException)
                                    {
                                        continue;
                                    }
                                    if (oledbConn == null) continue;

                                    string connString = oledbConn.Connection?.ToString() ?? "";
                                    bool isMashup = connString.Contains("Provider=Microsoft.Mashup.OleDb.1", StringComparison.OrdinalIgnoreCase);
                                    bool locationMatches = connString.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase);

                                    if (isMashup && locationMatches)
                                    {
                                        existingQueryTable = qTable;
                                        qTable = null; // Don't release - keeping reference
                                        targetWorksheet = worksheet;
                                        worksheet = null; // Don't release
                                        foundQueryTable = true;
                                        break;
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref oledbConn!);
                                    ComUtilities.Release(ref wbConn!);
                                    ComUtilities.Release(ref qTable!);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref queryTables!);
                        }

                        if (foundQueryTable) break;

                        // SECOND: Check for ListObject (Pattern 2 - from previous Update)
                        dynamic? listObjects = null;
                        try
                        {
                            listObjects = worksheet.ListObjects;
                            for (int lo = 1; lo <= listObjects.Count; lo++)
                            {
                                dynamic? listObj = null;
                                dynamic? queryTable = null;
                                dynamic? wbConn = null;
                                dynamic? oledbConn = null;
                                try
                                {
                                    listObj = listObjects.Item(lo);

                                    // NOTE: Accessing QueryTable on a regular Excel table (not from external data)
                                    // throws COMException 0x800A03EC. We must catch and skip such tables.
                                    try
                                    {
                                        queryTable = listObj.QueryTable;
                                    }
                                    catch (System.Runtime.InteropServices.COMException)
                                    {
                                        // Regular table without QueryTable - skip it
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
                                    bool locationMatches = connString.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase);

                                    if (isMashup && locationMatches)
                                    {
                                        existingQueryTable = queryTable;
                                        queryTable = null; // Don't release - keeping reference
                                        targetWorksheet = worksheet;
                                        worksheet = null; // Don't release
                                        break;
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref oledbConn!);
                                    ComUtilities.Release(ref wbConn!);
                                    ComUtilities.Release(ref queryTable!);
                                    ComUtilities.Release(ref listObj!);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref listObjects!);
                        }
                    }
                    finally
                    {
                        if (worksheet != null && targetWorksheet == null) ComUtilities.Release(ref worksheet!);
                    }

                    if (existingQueryTable != null) break;
                }

                // STEP 3: Update the M code with formatted version
                // Note: 0x800A03EC error can occur in certain workbook states (see Issue #323)
                // Retry doesn't help - it's a workbook state issue, not transient
                query.Formula = formattedMCode;

                // STEP 4: Refresh if requested
                if (refresh)
                {
                    if (existingQueryTable != null)
                    {
                        // Worksheet query: Use QueryTable.Refresh(false) for synchronous refresh
                        // This properly propagates column structure changes
                        existingQueryTable.Refresh(false);
                    }
                    else
                    {
                        // Data Model-only query (no worksheet table): Use connection.Refresh()
                        // Find the Power Query connection and refresh it
                        dynamic? connections = null;
                        try
                        {
                            connections = ctx.Book.Connections;
                            for (int i = 1; i <= connections.Count; i++)
                            {
                                dynamic? conn = null;
                                dynamic? oledbConn = null;
                                try
                                {
                                    conn = connections.Item(i);

                                    // NOTE: Accessing OLEDBConnection on non-OLEDB connection types
                                    // (e.g., Type=7 ThisWorkbookDataModel, Type=8 workbook connections)
                                    // throws COMException 0x800A03EC. We must catch and skip.
                                    try
                                    {
                                        oledbConn = conn.OLEDBConnection;
                                    }
                                    catch (System.Runtime.InteropServices.COMException)
                                    {
                                        continue;
                                    }
                                    if (oledbConn == null) continue;

                                    string connString = oledbConn.Connection?.ToString() ?? "";
                                    bool isMashup = connString.Contains("Provider=Microsoft.Mashup.OleDb.1", StringComparison.OrdinalIgnoreCase);
                                    bool locationMatches = connString.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase);

                                    if (isMashup && locationMatches)
                                    {
                                        conn.Refresh();
                                        break;
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref oledbConn!);
                                    ComUtilities.Release(ref conn!);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref connections!);
                        }
                    }
                }
                // Connection-only queries (no QueryTable, no Data Model connection) don't need refresh

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref existingQueryTable!);
                ComUtilities.Release(ref targetWorksheet!);
                ComUtilities.Release(ref worksheets!);
                ComUtilities.Release(ref query!);
                ComUtilities.Release(ref queries!);
            }
        });
    }

}


