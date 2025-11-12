using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.QueryTable;

/// <summary>
/// QueryTable management commands - FilePath-based API implementations
/// </summary>
public partial class QueryTableCommands
{
    /// <summary>
    /// Lists all QueryTables in the workbook with connection and range information
    /// </summary>
    public async Task<QueryTableListResult> ListAsync(string filePath)
    {
        var result = new QueryTableListResult { FilePath = filePath };
        var queryTables = new List<QueryTableInfo>();

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? worksheets = null;
                try
                {
                    worksheets = handle.Workbook.Worksheets;
                    for (int ws = 1; ws <= worksheets.Count; ws++)
                    {
                        dynamic? worksheet = null;
                        dynamic? sheetQueryTables = null;
                        try
                        {
                            worksheet = worksheets.Item(ws);
                            string worksheetName = worksheet.Name;
                            sheetQueryTables = worksheet.QueryTables;

                            for (int qt = 1; qt <= sheetQueryTables.Count; qt++)
                            {
                                dynamic? queryTable = null;
                                try
                                {
                                    queryTable = sheetQueryTables.Item(qt);
                                    var info = ExtractQueryTableInfo(queryTable, worksheetName);
                                    queryTables.Add(info);
                                }
                                finally
                                {
                                    ComUtilities.Release(ref queryTable);
                                }
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref sheetQueryTables);
                            ComUtilities.Release(ref worksheet);
                        }
                    }
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref worksheets);
                }
            });

            result.QueryTables = queryTables;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// Gets detailed information about a specific QueryTable
    /// </summary>
    public async Task<QueryTableInfoResult> GetAsync(string filePath, string queryTableName)
    {
        var result = new QueryTableInfoResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? queryTable = ComUtilities.FindQueryTable(handle.Workbook, queryTableName);
                if (queryTable == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"QueryTable '{queryTableName}' not found";
                    return;
                }

                try
                {
                    string worksheetName = FindWorksheetForQueryTable(handle.Workbook, queryTableName);
                    result.QueryTable = ExtractQueryTableInfo(queryTable, worksheetName);
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref queryTable);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// Deletes a QueryTable from the workbook
    /// </summary>
    public async Task<OperationResult> DeleteAsync(string filePath, string queryTableName)
    {
        var result = new OperationResult { FilePath = filePath, Action = "delete-querytable" };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? queryTable = ComUtilities.FindQueryTable(handle.Workbook, queryTableName);
                if (queryTable == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"QueryTable '{queryTableName}' not found";
                    return;
                }

                try
                {
                    queryTable.Delete();
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref queryTable);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    // Note: CreateFromConnectionAsync, CreateFromQueryAsync, RefreshAsync, RefreshAllAsync, UpdatePropertiesAsync
    // are complex operations that require ExcelWorkbookHandle.ExecuteAsync() pattern for proper COM handling.
    // These will be implemented after ExecuteAsync pattern is established for FilePath-based operations.
    // For now, they delegate to batch-based implementation (interim solution).
}
