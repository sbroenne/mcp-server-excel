using Sbroenne.ExcelMcp.Core.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
using Sbroenne.ExcelMcp.Core.Session;
using System.Text.RegularExpressions;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Excel Table (ListObject) management commands implementation
/// </summary>
public class TableCommands : ITableCommands
{
    /// <summary>
    /// Regular expression for valid table names (alphanumeric, underscore, no spaces, must start with letter or underscore)
    /// </summary>
    private static readonly Regex TableNameRegex = new(@"^[a-zA-Z_][a-zA-Z0-9_]*$", RegexOptions.Compiled);

    /// <summary>
    /// Maximum allowed table name length
    /// </summary>
    private const int MaxTableNameLength = 255;

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
                result.SuggestedNextActions.Add($"Use 'powerquery import' to reference table in Power Query: Excel.CurrentWorkbook(){{[Name=\"{tableName}\"]}}[Content]");
                result.SuggestedNextActions.Add($"Use 'table delete {tableName}' to remove table (converts back to range)");
                result.WorkflowHint = $"Table '{tableName}' created successfully. Ready for Power Query integration.";

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
                result.SuggestedNextActions.Add($"Update Power Query references to use new name: '{newName}'");
                result.WorkflowHint = $"Table renamed from '{tableName}' to '{newName}'. Update any Power Query references.";

                result.Success = true; return result;
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
                result.SuggestedNextActions.Add("Update Power Query expressions that referenced this table");
                result.WorkflowHint = $"Table '{tableName}' deleted. Data converted back to regular range.";

                result.Success = true; return result;
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
                result.SuggestedNextActions.Add($"Reference in Power Query: Excel.CurrentWorkbook(){{[Name=\"{tableName}\"]}}[Content]");
                result.WorkflowHint = $"Table '{tableName}' has {rowCount} rows and {columnCount} columns.";

                result.Success = true; return result;
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

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> ResizeAsync(IExcelBatch batch, string tableName, string newRange)
    {
// Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "resize" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? newRangeObj = null;
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
                newRangeObj = sheet.Range[newRange];

                // Resize the table
                table.Resize(newRangeObj);

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table info {tableName}' to verify the new size");
                result.SuggestedNextActions.Add("Use 'table read {tableName}' to read the updated data");
                result.WorkflowHint = $"Table '{tableName}' resized to {newRange}.";

                result.Success = true; return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref newRangeObj);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> ToggleTotalsAsync(IExcelBatch batch, string tableName, bool showTotals)
    {
// Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "toggle-totals" };
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

                table.ShowTotals = showTotals;

                result.Success = true;
                result.SuggestedNextActions.Add(showTotals
                    ? $"Use 'table set-column-total {tableName} <column> <function>' to configure totals"
                    : $"Use 'table toggle-totals {tableName} true' to re-enable totals");
                result.WorkflowHint = $"Totals row {(showTotals ? "enabled" : "disabled")} for table '{tableName}'.";

                result.Success = true; return result;
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
    public async Task<OperationResult> SetColumnTotalAsync(IExcelBatch batch, string tableName, string columnName, string totalFunction)
    {
// Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-column-total" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Ensure totals row is shown
                if (!table.ShowTotals)
                {
                    table.ShowTotals = true;
                }

                // Find the column
                listColumns = table.ListColumns;
                column = null;
                for (int i = 1; i <= listColumns.Count; i++)
                {
                    dynamic? col = null;
                    try
                    {
                        col = listColumns.Item(i);
                        if (col.Name == columnName)
                        {
                            column = col;
                            break;
                        }
                    }
                    finally
                    {
                        if (col != null && col.Name != columnName)
                        {
                            ComUtilities.Release(ref col);
                        }
                    }
                }

                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{columnName}' not found in table '{tableName}'";
                    return result;
                }

                // Map function name to Excel constant
                // xlTotalsCalculationSum = 1, xlTotalsCalculationAverage = 2, xlTotalsCalculationCount = 3,
                // xlTotalsCalculationCountNums = 4, xlTotalsCalculationMax = 5, xlTotalsCalculationMin = 6,
                // xlTotalsCalculationStdDev = 7, xlTotalsCalculationVar = 9, xlTotalsCalculationNone = 0
                int xlFunction = totalFunction.ToLowerInvariant() switch
                {
                    "sum" => 1,
                    "average" or "avg" => 2,
                    "count" => 3,
                    "countnums" => 4,
                    "max" => 5,
                    "min" => 6,
                    "stddev" => 7,
                    "var" => 9,
                    "none" => 0,
                    _ => throw new ArgumentException($"Unknown total function '{totalFunction}'. Valid: sum, average, count, countnums, max, min, stddev, var, none")
                };

                column.TotalsCalculation = xlFunction;

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table info {tableName}' to verify totals configuration");
                result.SuggestedNextActions.Add($"Use 'table read {tableName}' to see calculated totals");
                result.WorkflowHint = $"Column '{columnName}' total set to {totalFunction}.";

                result.Success = true; return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<TableDataResult> ReadDataAsync(IExcelBatch batch, string tableName)
    {
// Security: Validate table name
        ValidateTableName(tableName);

        var result = new TableDataResult { FilePath = batch.WorkbookPath, TableName = tableName };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? dataBodyRange = null;
            dynamic? headerRowRange = null;
            dynamic? listColumns = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Get headers
                if (table.ShowHeaders)
                {
                    listColumns = table.ListColumns;
                    for (int i = 1; i <= listColumns.Count; i++)
                    {
                        dynamic? column = null;
                        try
                        {
                            column = listColumns.Item(i);
                            result.Headers.Add(column.Name);
                        }
                        finally
                        {
                            ComUtilities.Release(ref column);
                        }
                    }
                }

                result.ColumnCount = table.ListColumns.Count;

                // Get data
                dataBodyRange = table.DataBodyRange;
                if (dataBodyRange != null)
                {
                    object[,] values = dataBodyRange.Value2;
                    if (values != null)
                    {
                        int rows = values.GetLength(0);
                        int cols = values.GetLength(1);
                        result.RowCount = rows;

                        for (int r = 1; r <= rows; r++)
                        {
                            var row = new List<object?>();
                            for (int c = 1; c <= cols; c++)
                            {
                                row.Add(values[r, c]);
                            }
                            result.Data.Add(row);
                        }
                    }
                }

                result.Success = true;
                result.SuggestedNextActions.Add("Use 'table append' to add more rows to this table");
                result.SuggestedNextActions.Add("Use worksheet 'write' to update the underlying range data");
                result.WorkflowHint = $"Read {result.RowCount} rows from table '{tableName}'.";

                result.Success = true; return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref headerRowRange);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> AppendRowsAsync(IExcelBatch batch, string tableName, string csvData)
    {
// Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "append-rows" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? dataBodyRange = null;
            dynamic? newRange = null;
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

                // Parse CSV data
                var lines = csvData.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0)
                {
                    result.Success = false;
                    result.ErrorMessage = "No data to append";
                    return result;
                }

                // Get current table size
                int currentRow;
                dataBodyRange = table.DataBodyRange;
                if (dataBodyRange != null)
                {
                    currentRow = dataBodyRange.Row + dataBodyRange.Rows.Count;
                }
                else
                {
                    // Table has only headers
                    dynamic? headerRange = null;
                    try
                    {
                        headerRange = table.HeaderRowRange;
                        currentRow = headerRange.Row + 1;
                    }
                    finally
                    {
                        ComUtilities.Release(ref headerRange);
                    }
                }

                int columnCount = table.ListColumns.Count;
                int rowsToAdd = lines.Length;

                // Write data to cells below the table
                for (int i = 0; i < lines.Length; i++)
                {
                    var values = lines[i].Split(',');
                    for (int j = 0; j < Math.Min(values.Length, columnCount); j++)
                    {
                        dynamic? cell = null;
                        try
                        {
                            cell = sheet.Cells[currentRow + i, table.Range.Column + j];
                            cell.Value2 = values[j].Trim().Trim('"');
                        }
                        finally
                        {
                            ComUtilities.Release(ref cell);
                        }
                    }
                }

                // Resize table to include new rows
                int newLastRow = currentRow + rowsToAdd - 1;
                int newLastCol = table.Range.Column + columnCount - 1;
                string newRangeAddress = $"{sheet.Cells[table.Range.Row, table.Range.Column].Address}:{sheet.Cells[newLastRow, newLastCol].Address}";

                dynamic? resizeRange = null;
                try
                {
                    resizeRange = sheet.Range[newRangeAddress];
                    table.Resize(resizeRange);
                }
                finally
                {
                    ComUtilities.Release(ref resizeRange);
                }

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table read {tableName}' to verify appended data");
                result.SuggestedNextActions.Add($"Use 'table info {tableName}' to see updated row count");
                result.WorkflowHint = $"Appended {rowsToAdd} rows to table '{tableName}'. Table auto-expanded.";

                result.Success = true; return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref newRange);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
#pragma warning disable CS1998 // Async method lacks await operators (synchronous COM interop)
    public async Task<OperationResult> SetStyleAsync(IExcelBatch batch, string tableName, string tableStyle)
    {
// Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-style" };
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

                table.TableStyle = tableStyle;

                result.Success = true;
                result.SuggestedNextActions.Add($"Use 'table info {tableName}' to verify the style change");
                result.SuggestedNextActions.Add("Common styles: TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11");
                result.WorkflowHint = $"Table '{tableName}' style changed to '{tableStyle}'.";

                result.Success = true; return result;
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
    public async Task<OperationResult> AddToDataModelAsync(IExcelBatch batch, string tableName)
    {
// Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "add-to-data-model" };
        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? modelTables = null;
            dynamic? connections = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Check if workbook has a Data Model (Model object)
                dynamic? model = null;
                try
                {
                    model = ctx.Book.Model;
                    if (model == null)
                    {
                        result.Success = false;
                        result.ErrorMessage = "Workbook does not have a Data Model. Data Model is only available in Excel 2013+ with Power Pivot enabled.";
                        return result;
                    }
                }
                catch
                {
                    result.Success = false;
                    result.ErrorMessage = "Data Model not available. Ensure Excel has Power Pivot add-in enabled.";
                    return result;
                }

                // Check if table is already in the Data Model
                try
                {
                    modelTables = model.ModelTables;
                    for (int i = 1; i <= modelTables.Count; i++)
                    {
                        dynamic? modelTable = null;
                        try
                        {
                            modelTable = modelTables.Item(i);
                            string sourceTableName = modelTable.SourceName;
                            if (sourceTableName == tableName || sourceTableName.EndsWith($"[{tableName}]"))
                            {
                                result.Success = false;
                                result.ErrorMessage = $"Table '{tableName}' is already in the Data Model";
                                return result;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref modelTable);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelTables);
                }

                // Create a connection for the table
                string connectionName = $"WorkbookConnection_{tableName}";
                string connectionString = $"WORKSHEET;{ctx.Book.FullName}";
                string commandText = $"SELECT * FROM [{tableName}]";

                // Check if connection already exists
                connections = ctx.Book.Connections;
                bool connectionExists = false;
                for (int i = 1; i <= connections.Count; i++)
                {
                    dynamic? conn = null;
                    try
                    {
                        conn = connections.Item(i);
                        if (conn.Name == connectionName)
                        {
                            connectionExists = true;
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref conn);
                    }
                }

                // Add table to Data Model
                // Using numeric constant for xlConnectionTypeOLEDB = 3
                if (!connectionExists)
                {
                    try
                    {
                        dynamic? newConnection = connections.Add2(
                            connectionName,
                            "Connection to Excel Table",
                            connectionString,
                            commandText,
                            3, // xlConnectionTypeOLEDB
                            true, // SSO (not used for local)
                            false // AddToModel parameter
                        );
                        ComUtilities.Release(ref newConnection);
                    }
                    catch
                    {
                        // Connection might not be needed in some Excel versions
                        // Continue anyway
                    }
                }

                // Add the table to the model using ModelTables.Add
                try
                {
                    modelTables = model.ModelTables;
                    dynamic? newModelTable = modelTables.Add(
                        connectionName,
                        tableName
                    );
                    ComUtilities.Release(ref newModelTable);
                    ComUtilities.Release(ref modelTables);
                }
                catch (Exception ex)
                {
                    // Try alternative approach: use Publish to Data Model
                    try
                    {
                        // Some Excel versions support PublishToDataModel on ListObject
                        table.Publish(null, false); // Publish to Data Model
                    }
                    catch
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Failed to add table to Data Model: {ex.Message}. Ensure Power Pivot is enabled.";
                        return result;
                    }
                }

                ComUtilities.Release(ref model);

                result.Success = true;
                result.SuggestedNextActions.Add("Use 'dm-list-tables' to verify the table is in the Data Model");
                result.SuggestedNextActions.Add($"Use 'dm-create-measure' to add DAX measures based on '{tableName}'");
                result.SuggestedNextActions.Add("Use 'dm-refresh' to refresh the Data Model");
                result.WorkflowHint = $"Table '{tableName}' added to Power Pivot Data Model.";

                result.Success = true; return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref connections);
                ComUtilities.Release(ref modelTables);
                ComUtilities.Release(ref table);
            }
        });
    }

    #region Private Helper Methods

    /// <summary>
    /// Validates a table name to prevent injection attacks and ensure Excel compatibility
    /// </summary>
    /// <param name="tableName">Table name to validate</param>
    /// <exception cref="ArgumentException">Thrown if table name is invalid</exception>
    private static void ValidateTableName(string tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Table name cannot be null or empty", nameof(tableName));
        }

        if (tableName.Length > MaxTableNameLength)
        {
            throw new ArgumentException(
                $"Table name too long: {tableName.Length} characters (maximum: {MaxTableNameLength})",
                nameof(tableName));
        }

        if (!TableNameRegex.IsMatch(tableName))
        {
            throw new ArgumentException(
                $"Invalid table name '{tableName}'. Table names must start with a letter or underscore, " +
                "and can only contain letters, numbers, and underscores (no spaces or special characters).",
                nameof(tableName));
        }

        // Check for reserved names
        string upperName = tableName.ToUpperInvariant();
        if (upperName == "PRINT_AREA" || upperName == "PRINT_TITLES" ||
            upperName == "_XLNM" || upperName.StartsWith("_XLNM."))
        {
            throw new ArgumentException(
                $"Table name '{tableName}' is reserved by Excel",
                nameof(tableName));
        }
    }

    /// <summary>
    /// Finds a table by name in the workbook
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to find</param>
    /// <returns>The table object if found, null otherwise</returns>
    private static dynamic? FindTable(dynamic workbook, string tableName)
    {
        dynamic? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            for (int i = 1; i <= sheets.Count; i++)
            {
                dynamic? sheet = null;
                dynamic? listObjects = null;
                try
                {
                    sheet = sheets.Item(i);
                    listObjects = sheet.ListObjects;

                    for (int j = 1; j <= listObjects.Count; j++)
                    {
                        dynamic? table = null;
                        try
                        {
                            table = listObjects.Item(j);
                            if (table.Name == tableName)
                            {
                                // Found it - return without releasing
                                return table;
                            }
                        }
                        finally
                        {
                            if (table != null && table.Name != tableName)
                            {
                                // Only release if not returning this table
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref listObjects);
                    ComUtilities.Release(ref sheet);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }

        return null;
    }

    /// <summary>
    /// Checks if a table with the given name exists in the workbook
    /// </summary>
    /// <param name="workbook">The workbook to search</param>
    /// <param name="tableName">Name of the table to check</param>
    /// <returns>True if table exists, false otherwise</returns>
    private static bool TableExists(dynamic workbook, string tableName)
    {
        dynamic? table = FindTable(workbook, tableName);
        bool exists = table != null;
        ComUtilities.Release(ref table);
        return exists;
    }

    #endregion
}






