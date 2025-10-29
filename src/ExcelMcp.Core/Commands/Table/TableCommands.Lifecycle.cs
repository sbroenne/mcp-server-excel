using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table lifecycle operations (List, Create, Rename, Delete, GetInfo)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<TableListResult> ListAsync(IExcelBatch batch)
    {
        var result = new TableListResult { FilePath = batch.WorkbookPath };
        
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheets = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    dynamic? listObjects = null;
                    try
                    {
                        sheet = sheets.Item(i);
                        listObjects = sheet.ListObjects;
                        string sheetName = sheet.Name;

                        for (int j = 1; j <= listObjects.Count; j++)
                        {
                            dynamic? table = null;
                            dynamic? headerRowRange = null;
                            dynamic? dataBodyRange = null;
                            try
                            {
                                table = listObjects.Item(j);
                                string tableName = table.Name;
                                string rangeAddress = table.Range.Address;
                                bool showHeaders = table.ShowHeaders;
                                bool showTotals = table.ShowTotals;
                                string tableStyleName = table.TableStyle?.Name ?? "";

                                // Get column count and names
                                int columnCount = table.ListColumns.Count;
                                var columns = new List<string>();

                                if (showHeaders)
                                {
                                    dynamic? listColumns = null;
                                    try
                                    {
                                        listColumns = table.ListColumns;
                                        for (int k = 1; k <= listColumns.Count; k++)
                                        {
                                            dynamic? column = null;
                                            try
                                            {
                                                column = listColumns.Item(k);
                                                columns.Add(column.Name);
                                            }
                                            finally
                                            {
                                                ComUtilities.Release(ref column);
                                            }
                                        }
                                    }
                                    finally
                                    {
                                        ComUtilities.Release(ref listColumns);
                                    }
                                }

                                // Get row count (excluding header)
                                // SECURITY FIX: DataBodyRange can be NULL if table has only headers
                                int rowCount = 0;
                                try
                                {
                                    dataBodyRange = table.DataBodyRange;
                                    if (dataBodyRange != null)
                                    {
                                        rowCount = dataBodyRange.Rows.Count;
                                    }
                                }
                                finally
                                {
                                    ComUtilities.Release(ref dataBodyRange);
                                }

                                result.Tables.Add(new TableInfo
                                {
                                    Name = tableName,
                                    SheetName = sheetName,
                                    Range = rangeAddress,
                                    HasHeaders = showHeaders,
                                    TableStyle = tableStyleName,
                                    RowCount = rowCount,
                                    ColumnCount = columnCount,
                                    Columns = columns,
                                    ShowTotals = showTotals
                                });
                            }
                            finally
                            {
                                ComUtilities.Release(ref headerRowRange);
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
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> CreateAsync(IExcelBatch batch, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "create" };
        
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? rangeObj = null;
            dynamic? listObjects = null;
            dynamic? newTable = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Sheet '{sheetName}' not found";
                    return result;
                }

                // Check if table name already exists
                if (TableExists(ctx.Book, tableName))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' already exists";
                    return result;
                }

                // Get the range to convert to table
                rangeObj = sheet.Range[range];

                listObjects = sheet.ListObjects;

                // Create table using numeric constant (xlSrcRange = 1)
                // XlListObjectSourceType.xlSrcRange causes enum assembly loading issues
                int xlSrcRange = 1;
                int xlYes = 1;  // xlYes for has headers
                int xlGuess = 0;  // xlGuess
                int headerOption = hasHeaders ? xlYes : xlGuess;

                newTable = listObjects.Add(xlSrcRange, rangeObj, null, headerOption);
                newTable.Name = tableName;

                // Apply table style if specified
                if (!string.IsNullOrWhiteSpace(tableStyle))
                {
                    newTable.TableStyle = tableStyle;
                }

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table info {tableName}' to view table details");
                result.SuggestedNextActions.Add($"Use structured references in formulas: ={tableName}[@Column] or =[@Column] within table");
                result.SuggestedNextActions.Add($"Use 'table delete {tableName}' to remove table (converts back to range)");
                result.WorkflowHint = $"Table '{tableName}' created successfully. AutoFilter, structured references, and dynamic expansion enabled.";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref newTable);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref rangeObj);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> RenameAsync(IExcelBatch batch, string tableName, string newName)
    {
        // Security: Validate table names
        ValidateTableName(tableName);
        ValidateTableName(newName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "rename" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Check if new name already exists
                if (TableExists(ctx.Book, newName))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{newName}' already exists";
                    return result;
                }

                table.Name = newName;
                result.Success = true;
                result.SuggestedNextActions.Add($"Update structured references in formulas to use new name: ={newName}[@Column]");
                result.WorkflowHint = $"Table renamed from '{tableName}' to '{newName}'. Update formulas using structured references.";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> DeleteAsync(IExcelBatch batch, string tableName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "delete" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? tableRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // SECURITY FIX: Store range info before Unlist() for proper cleanup
                try
                {
                    tableRange = table.Range;
                }
                catch
                {
                    // Ignore if range is not accessible
                }

                // Convert table back to range (Unlist)
                table.Unlist();

                // SECURITY FIX: After Unlist(), we must explicitly release the table COM object
                // The table object is no longer valid but still holds a COM reference
                ComUtilities.Release(ref table);

                result.Success = true;
                result.SuggestedNextActions.Add("Data remains in worksheet as a regular range");
                result.SuggestedNextActions.Add("Update formulas that used structured references to this table");
                result.WorkflowHint = $"Table '{tableName}' deleted. Data converted back to regular range.";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<TableInfoResult> GetInfoAsync(IExcelBatch batch, string tableName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new TableInfoResult { FilePath = batch.WorkbookPath };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? dataBodyRange = null;
            dynamic? headerRowRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                sheet = table.Parent;
                string sheetName = sheet.Name;
                string rangeAddress = table.Range.Address;
                bool showHeaders = table.ShowHeaders;
                bool showTotals = table.ShowTotals;
                string tableStyleName = table.TableStyle?.Name ?? "";

                // Get column count and names
                int columnCount = table.ListColumns.Count;
                var columns = new List<string>();

                if (showHeaders)
                {
                    dynamic? listColumns = null;
                    try
                    {
                        listColumns = table.ListColumns;
                        for (int i = 1; i <= listColumns.Count; i++)
                        {
                            dynamic? column = null;
                            try
                            {
                                column = listColumns.Item(i);
                                columns.Add(column.Name);
                            }
                            finally
                            {
                                ComUtilities.Release(ref column);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref listColumns);
                    }
                }

                // Get row count (excluding header)
                // SECURITY FIX: DataBodyRange can be NULL if table has only headers
                int rowCount = 0;
                try
                {
                    dataBodyRange = table.DataBodyRange;
                    if (dataBodyRange != null)
                    {
                        rowCount = dataBodyRange.Rows.Count;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref dataBodyRange);
                }

                result.Table = new TableInfo
                {
                    Name = tableName,
                    SheetName = sheetName,
                    Range = rangeAddress,
                    HasHeaders = showHeaders,
                    TableStyle = tableStyleName,
                    RowCount = rowCount,
                    ColumnCount = columnCount,
                    Columns = columns,
                    ShowTotals = showTotals
                };

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table rename {tableName} NewName' to rename table");
                result.SuggestedNextActions.Add($"Use 'table delete {tableName}' to remove table");
                result.SuggestedNextActions.Add($"Use structured references in formulas: ={tableName}[@Column]");
                result.WorkflowHint = $"Table '{tableName}' has {rowCount} rows and {columnCount} columns.";

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref headerRowRange);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }
#pragma warning restore CS1998
}
