using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

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
    /// UPDATED to check for BOTH QueryTable (from LoadTo) and ListObject (from previous Update) patterns.
    /// </summary>
    public OperationResult Update(IExcelBatch batch, string queryName, string mCode)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "update"
        };

        if (!ValidateQueryName(queryName, out string? validationError))
        {
            result.Success = false;
            result.ErrorMessage = validationError;
            return result;
        }

        if (string.IsNullOrWhiteSpace(mCode))
        {
            result.Success = false;
            result.ErrorMessage = "M code cannot be empty";
            return result;
        }

        try
        {
            return batch.Execute((ctx, ct) =>
            {
                dynamic? queries = null;
                dynamic? query = null;
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
                            if (q != null) ComUtilities.Release(ref q!);
                        }
                    }

                    if (query == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Query '{queryName}' not found";
                        return result;
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

                                        oledbConn = wbConn.OLEDBConnection;
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
                                        if (oledbConn != null) ComUtilities.Release(ref oledbConn!);
                                        if (wbConn != null) ComUtilities.Release(ref wbConn!);
                                        if (qTable != null) ComUtilities.Release(ref qTable!);
                                    }
                                }
                            }
                            finally
                            {
                                if (queryTables != null) ComUtilities.Release(ref queryTables!);
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
                                        queryTable = listObj.QueryTable;
                                        if (queryTable == null) continue;

                                        wbConn = queryTable.WorkbookConnection;
                                        if (wbConn == null) continue;

                                        oledbConn = wbConn.OLEDBConnection;
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
                                        if (oledbConn != null) ComUtilities.Release(ref oledbConn!);
                                        if (wbConn != null) ComUtilities.Release(ref wbConn!);
                                        if (queryTable != null) ComUtilities.Release(ref queryTable!);
                                        if (listObj != null) ComUtilities.Release(ref listObj!);
                                    }
                                }
                            }
                            finally
                            {
                                if (listObjects != null) ComUtilities.Release(ref listObjects!);
                            }
                        }
                        finally
                        {
                            if (worksheet != null && targetWorksheet == null) ComUtilities.Release(ref worksheet!);
                        }

                        if (existingQueryTable != null) break;
                    }

                    // STEP 3: Update the M code
                    query.Formula = mCode;

                    // STEP 4: Refresh existing QueryTable if it exists
                    if (existingQueryTable != null)
                    {
                        try
                        {
                            // Just refresh - PreserveColumnInfo=false allows schema changes
                            existingQueryTable.Refresh(false);  // Synchronous
                            result.Success = true;
                            result.Action = "updated and refreshed";
                        }
                        catch (Exception ex)
                        {
                            result.Success = false;
                            result.ErrorMessage = $"Failed to refresh after update: {ex.Message}";
                        }
                    }
                    else
                    {
                        // No QueryTable or ListObject - connection-only query
                        result.Success = true;
                        result.Action = "updated (connection-only)";
                    }

                    return result;
                }
                finally
                {
                    if (existingQueryTable != null) ComUtilities.Release(ref existingQueryTable!);
                    if (targetWorksheet != null) ComUtilities.Release(ref targetWorksheet!);
                    if (worksheets != null) ComUtilities.Release(ref worksheets!);
                    if (query != null) ComUtilities.Release(ref query!);
                    if (queries != null) ComUtilities.Release(ref queries!);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Unexpected error: {ex.Message}";
            return result;
        }
    }

}
