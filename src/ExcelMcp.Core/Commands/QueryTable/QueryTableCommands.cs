using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.QueryTable;

/// <summary>
/// QueryTable management commands - Core data layer (no console output)
/// Leverages existing PowerQueryHelpers infrastructure for QueryTable operations
/// </summary>
public partial class QueryTableCommands : IQueryTableCommands
{
    /// <inheritdoc />
    public async Task<QueryTableListResult> ListAsync(IExcelBatch batch)
    {
        return await batch.Execute((ctx, ct) =>
        {
            var result = new QueryTableListResult { FilePath = batch.WorkbookPath };
            var queryTables = new List<QueryTableInfo>();

            // Iterate through all worksheets and their QueryTables
            dynamic? worksheets = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
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
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
            }

            result.QueryTables = queryTables;
            result.Success = true;
            return result;
        });
    }

    /// <inheritdoc />
    public async Task<QueryTableInfoResult> GetAsync(IExcelBatch batch, string queryTableName)
    {
        return await batch.Execute((ctx, ct) =>
        {
            var result = new QueryTableInfoResult { FilePath = batch.WorkbookPath };

            dynamic? queryTable = ComUtilities.FindQueryTable(ctx.Book, queryTableName);
            if (queryTable == null)
            {
                result.Success = false;
                result.ErrorMessage = $"QueryTable '{queryTableName}' not found";
                return result;
            }

            try
            {
                // Find the worksheet containing this QueryTable
                string worksheetName = FindWorksheetForQueryTable(ctx.Book, queryTableName);
                result.QueryTable = ExtractQueryTableInfo(queryTable, worksheetName);
                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAsync(IExcelBatch batch, string queryTableName, TimeSpan? timeout = null)
    {
        return await batch.Execute((ctx, ct) =>
        {
            var result = new OperationResult
            {
                FilePath = batch.WorkbookPath,
                Action = "refresh"
            };

            dynamic? queryTable = ComUtilities.FindQueryTable(ctx.Book, queryTableName);
            if (queryTable == null)
            {
                result.Success = false;
                result.ErrorMessage = $"QueryTable '{queryTableName}' not found";
                return result;
            }

            try
            {
                // CRITICAL: Use synchronous refresh for persistence
                queryTable.Refresh(false);
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to refresh QueryTable: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }

            return result;
        }, timeout: timeout);
    }

    /// <inheritdoc />
    public async Task<OperationResult> RefreshAllAsync(IExcelBatch batch, TimeSpan? timeout = null)
    {
        return await batch.Execute((ctx, ct) =>
        {
            var result = new OperationResult
            {
                FilePath = batch.WorkbookPath,
                Action = "refresh-all"
            };

            int refreshedCount = 0;
            var errors = new List<string>();

            dynamic? worksheets = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                for (int ws = 1; ws <= worksheets.Count; ws++)
                {
                    dynamic? worksheet = null;
                    dynamic? sheetQueryTables = null;
                    try
                    {
                        worksheet = worksheets.Item(ws);
                        sheetQueryTables = worksheet.QueryTables;

                        for (int qt = 1; qt <= sheetQueryTables.Count; qt++)
                        {
                            dynamic? queryTable = null;
                            try
                            {
                                queryTable = sheetQueryTables.Item(qt);
                                string queryTableName = queryTable.Name?.ToString() ?? "";

                                try
                                {
                                    // CRITICAL: Use synchronous refresh for each QueryTable
                                    queryTable.Refresh(false);
                                    refreshedCount++;
                                }
                                catch (Exception ex)
                                {
                                    errors.Add($"{queryTableName}: {ex.Message}");
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
                        ComUtilities.Release(ref sheetQueryTables);
                        ComUtilities.Release(ref worksheet);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref worksheets);
            }

            if (errors.Count > 0)
            {
                result.Success = false;
                result.ErrorMessage = $"Refreshed {refreshedCount} QueryTable(s), but {errors.Count} failed: {string.Join("; ", errors)}";
            }
            else
            {
                result.Success = true;
                result.OperationContext = new Dictionary<string, object>
                {
                    { "RefreshedCount", refreshedCount }
                };
            }

            return result;
        }, timeout: timeout);
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string queryTableName)
    {
        return await batch.Execute((ctx, ct) =>
        {
            var result = new OperationResult
            {
                FilePath = batch.WorkbookPath,
                Action = "delete"
            };

            dynamic? queryTable = ComUtilities.FindQueryTable(ctx.Book, queryTableName);
            if (queryTable == null)
            {
                result.Success = false;
                result.ErrorMessage = $"QueryTable '{queryTableName}' not found";
                return result;
            }

            try
            {
                queryTable.Delete();
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to delete QueryTable: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref queryTable);
            }

            return result;
        });
    }

    #region Helper Methods

    /// <summary>
    /// Extracts information from a QueryTable COM object
    /// </summary>
    private static QueryTableInfo ExtractQueryTableInfo(dynamic queryTable, string worksheetName)
    {
        var info = new QueryTableInfo
        {
            WorksheetName = worksheetName,
            Name = queryTable.Name?.ToString() ?? "",
            ConnectionString = queryTable.Connection?.ToString() ?? "",
            CommandText = queryTable.CommandText?.ToString() ?? "",
            BackgroundQuery = TryGetBool(queryTable, "BackgroundQuery"),
            RefreshOnFileOpen = TryGetBool(queryTable, "RefreshOnFileOpen"),
            PreserveColumnInfo = TryGetBool(queryTable, "PreserveColumnInfo"),
            PreserveFormatting = TryGetBool(queryTable, "PreserveFormatting"),
            AdjustColumnWidth = TryGetBool(queryTable, "AdjustColumnWidth")
        };

        // Extract range information
        try
        {
            dynamic? resultRange = queryTable.ResultRange;
            if (resultRange != null)
            {
                try
                {
                    info.Range = resultRange.Address?.ToString() ?? "";
                    info.RowCount = Convert.ToInt32(resultRange.Rows.Count);
                    info.ColumnCount = Convert.ToInt32(resultRange.Columns.Count);
                }
                finally
                {
                    ComUtilities.Release(ref resultRange);
                }
            }
        }
        catch
        {
            // Ignore errors getting range info
        }

        // Extract last refresh time
        try
        {
            var refreshDate = queryTable.RefreshDate;
            if (refreshDate != null)
            {
                if (refreshDate is DateTime dt)
                {
                    info.LastRefresh = dt;
                }
                else if (refreshDate is double dbl)
                {
                    info.LastRefresh = DateTime.FromOADate(dbl);
                }
            }
        }
        catch
        {
            // Ignore errors getting refresh date
        }

        return info;
    }

    /// <summary>
    /// Safely gets a boolean property from a QueryTable COM object
    /// </summary>
    private static bool TryGetBool(dynamic obj, string propertyName)
    {
        try
        {
            return propertyName switch
            {
                "BackgroundQuery" => obj.BackgroundQuery,
                "RefreshOnFileOpen" => obj.RefreshOnFileOpen,
                "PreserveColumnInfo" => obj.PreserveColumnInfo,
                "PreserveFormatting" => obj.PreserveFormatting,
                "AdjustColumnWidth" => obj.AdjustColumnWidth,
                _ => false
            };
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Finds the worksheet name containing a QueryTable
    /// </summary>
    private static string FindWorksheetForQueryTable(dynamic workbook, string queryTableName)
    {
        dynamic? worksheets = null;
        try
        {
            worksheets = workbook.Worksheets;
            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic? worksheet = null;
                dynamic? sheetQueryTables = null;
                try
                {
                    worksheet = worksheets.Item(ws);
                    sheetQueryTables = worksheet.QueryTables;

                    for (int qt = 1; qt <= sheetQueryTables.Count; qt++)
                    {
                        dynamic? queryTable = null;
                        try
                        {
                            queryTable = sheetQueryTables.Item(qt);
                            string currentName = queryTable.Name?.ToString() ?? "";

                            if (currentName.Equals(queryTableName, StringComparison.OrdinalIgnoreCase))
                            {
                                return worksheet.Name;
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
                    ComUtilities.Release(ref sheetQueryTables);
                    ComUtilities.Release(ref worksheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref worksheets);
        }

        return string.Empty;
    }

    #endregion
}
