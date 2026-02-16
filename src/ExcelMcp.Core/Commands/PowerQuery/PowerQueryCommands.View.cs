using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query View operations - STANDALONE implementation.
/// Based on Microsoft WorkbookQuery object model documentation.
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// View Power Query details: M code, description, and load configuration.
    /// STANDALONE implementation following Microsoft WorkbookQuery API.
    /// </summary>
    /// <remarks>
    /// Microsoft Docs Reference:
    /// - WorkbookQuery.Name property (Read/Write String)
    /// - WorkbookQuery.Description property (Read/Write String)
    /// - WorkbookQuery.Formula property (Read/Write String) - The Power Query M code
    ///
    /// Load configuration detection follows the pattern established in Update:
    /// - QueryTable (from LoadTo/Create) - created via sheet.QueryTables.Add()
    /// - ListObject (from previous Update) - created via sheet.ListObjects.Add()
    /// Both are checked to determine if query is connection-only or loaded to worksheet.
    /// </remarks>
    public PowerQueryViewResult View(IExcelBatch batch, string queryName)
    {
        var result = new PowerQueryViewResult
        {
            FilePath = batch.WorkbookPath,
            QueryName = queryName
        };

        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? queries = null;
            dynamic? query = null;
            dynamic? worksheets = null;

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
                            q = null; // Don't release - keeping reference
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

                // STEP 2: Read WorkbookQuery properties (per Microsoft docs)
                string mCode = query.Formula?.ToString() ?? "";
                result.MCode = mCode;
                result.CharacterCount = mCode.Length;

                // STEP 3: Detect load configuration (QueryTable or ListObject pattern)
                // Same detection logic as Update() - check BOTH patterns
                bool isLoadedToWorksheet = false;

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

                                    // Non-OLEDB connection types (Type=7, Type=8) throw COMException
                                    try { oledbConn = wbConn.OLEDBConnection; }
                                    catch (System.Runtime.InteropServices.COMException) { continue; }
                                    if (oledbConn == null) continue;

                                    string connString = oledbConn.Connection?.ToString() ?? "";
                                    bool isMashup = connString.Contains("Provider=Microsoft.Mashup.OleDb.1", StringComparison.OrdinalIgnoreCase);
                                    bool locationMatches = connString.Contains($"Location={queryName}", StringComparison.OrdinalIgnoreCase);

                                    if (isMashup && locationMatches)
                                    {
                                        isLoadedToWorksheet = true;
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

                        if (isLoadedToWorksheet) break;

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
                                        isLoadedToWorksheet = true;
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
                        ComUtilities.Release(ref worksheet!);
                    }

                    if (isLoadedToWorksheet) break;
                }

                // STEP 4: Set load mode result
                result.IsConnectionOnly = !isLoadedToWorksheet;

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref worksheets!);
                ComUtilities.Release(ref query!);
                ComUtilities.Release(ref queries!);
            }
        });
    }
}


