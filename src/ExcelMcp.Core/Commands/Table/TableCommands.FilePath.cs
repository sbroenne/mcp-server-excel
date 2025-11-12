using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// FilePath-based Table commands (file handle manager pattern)
/// </summary>
public partial class TableCommands
{
    /// <summary>
    /// Lists all Excel Tables in the workbook (filePath-based)
    /// </summary>
    public async Task<TableListResult> ListAsync(string filePath)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new TableListResult { FilePath = filePath };

        dynamic? sheets = null;
        try
        {
            sheets = handle.Workbook.Worksheets;
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
                        dynamic? dataBodyRange = null;
                        try
                        {
                            table = listObjects.Item(j);
                            string tableName = table.Name;
                            string rangeAddress = table.Range.Address;
                            bool showHeaders = table.ShowHeaders;
                            bool showTotals = table.ShowTotals;
                            string tableStyleName = table.TableStyle?.Name ?? "";

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
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error listing tables: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }

        return result;
    }

    /// <summary>
    /// Creates a new Excel Table from a range (filePath-based)
    /// </summary>
    public async Task<OperationResult> CreateAsync(string filePath, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? sheet = null;
        dynamic? targetRange = null;
        dynamic? table = null;
        try
        {
            sheet = handle.Workbook.Worksheets.Item(sheetName);
            targetRange = sheet.Range[range];

            dynamic? listObjects = null;
            try
            {
                listObjects = sheet.ListObjects;
                table = listObjects.Add(1, targetRange, Type.Missing, hasHeaders ? 1 : 2);
                table.Name = tableName;

                if (!string.IsNullOrEmpty(tableStyle))
                {
                    table.TableStyle = tableStyle;
                }

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref listObjects);
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error creating table: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref targetRange);
            ComUtilities.Release(ref sheet);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Renames an Excel Table (filePath-based)
    /// </summary>
    public async Task<OperationResult> RenameAsync(string filePath, string tableName, string newName)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? table = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            table.Name = newName;
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error renaming table: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref table);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Deletes an Excel Table (filePath-based)
    /// </summary>
    public async Task<OperationResult> DeleteAsync(string filePath, string tableName)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? table = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            table.Unlist();
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error deleting table: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref table);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Gets detailed information about an Excel Table (filePath-based)
    /// </summary>
    public async Task<TableInfoResult> GetAsync(string filePath, string tableName)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new TableInfoResult { FilePath = filePath };

        dynamic? table = null;
        dynamic? dataBodyRange = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            string sheetName = table.Parent.Name;
            string range = table.Range.Address;
            bool hasHeaders = table.ShowHeaders;
            bool showTotals = table.ShowTotals;
            string tableStyle = table.TableStyle?.Name ?? "";

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

            int columnCount = table.ListColumns.Count;

            var columns = new List<string>();
            if (hasHeaders)
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

            result.Table = new TableInfo
            {
                Name = tableName,
                SheetName = sheetName,
                Range = range,
                HasHeaders = hasHeaders,
                TableStyle = tableStyle,
                RowCount = rowCount,
                ColumnCount = columnCount,
                Columns = columns,
                ShowTotals = showTotals
            };
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error getting table info: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref table);
        }

        return result;
    }

    /// <summary>
    /// Resizes an Excel Table to a new range (filePath-based)
    /// </summary>
    public async Task<OperationResult> ResizeAsync(string filePath, string tableName, string newRange)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? table = null;
        dynamic? targetRange = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            dynamic? sheet = null;
            try
            {
                sheet = table.Parent;
                targetRange = sheet.Range[newRange];
                table.Resize(targetRange);
                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error resizing table: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref targetRange);
            ComUtilities.Release(ref table);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Toggles the totals row for an Excel Table (filePath-based)
    /// </summary>
    public async Task<OperationResult> ToggleTotalsAsync(string filePath, string tableName, bool showTotals)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? table = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            table.ShowTotals = showTotals;
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error toggling totals: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref table);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Sets the totals function for a specific column (filePath-based)
    /// </summary>
    public async Task<OperationResult> SetColumnTotalAsync(string filePath, string tableName, string columnName, string totalFunction)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? table = null;
        dynamic? column = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            if (!table.ShowTotals)
            {
                result.Success = false;
                result.ErrorMessage = "Totals row is not enabled";
                return result;
            }

            column = FindTableColumn(table, columnName);
            if (column == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Column '{columnName}' not found";
                return result;
            }

            int xlFunction = ParseTotalFunction(totalFunction);
            column.TotalsCalculation = xlFunction;
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error setting column total: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref column);
            ComUtilities.Release(ref table);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Appends rows to an Excel Table (filePath-based)
    /// </summary>
    public async Task<OperationResult> AppendAsync(string filePath, string tableName, List<List<object?>> rows)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? table = null;
        dynamic? dataBodyRange = null;
        dynamic? newRow = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            int columnCount = table.ListColumns.Count;

            foreach (var row in rows)
            {
                if (row.Count != columnCount)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Row has {row.Count} values but table has {columnCount} columns";
                    return result;
                }

                dynamic? listRow = null;
                try
                {
                    listRow = table.ListRows.Add();
                    newRow = listRow.Range;

                    for (int i = 0; i < row.Count; i++)
                    {
                        dynamic? cell = null;
                        try
                        {
                            cell = newRow.Cells[1, i + 1];
                            cell.Value2 = row[i];
                        }
                        finally
                        {
                            ComUtilities.Release(ref cell);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref newRow);
                    ComUtilities.Release(ref listRow);
                }
            }

            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error appending rows: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref dataBodyRange);
            ComUtilities.Release(ref table);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Changes the style of an Excel Table (filePath-based)
    /// </summary>
    public async Task<OperationResult> SetStyleAsync(string filePath, string tableName, string tableStyle)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? table = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            table.TableStyle = tableStyle;
            result.Success = true;
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error setting table style: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref table);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Adds an Excel Table to the Power Pivot Data Model (filePath-based)
    /// </summary>
    public async Task<OperationResult> AddToDataModelAsync(string filePath, string tableName)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = new OperationResult();

        dynamic? table = null;
        try
        {
            table = FindTable(handle.Workbook, tableName);
            if (table == null)
            {
                result.Success = false;
                result.ErrorMessage = $"Table '{tableName}' not found";
                return result;
            }

            // Check if already in data model
            dynamic? connection = null;
            try
            {
                connection = handle.Workbook.Connections.Item(tableName);
                if (connection != null)
                {
                    result.Success = false;
                    result.ErrorMessage = "Table is already in the Data Model";
                    return result;
                }
            }
            catch
            {
                // Not in data model yet, continue
            }
            finally
            {
                ComUtilities.Release(ref connection);
            }

            // Add to data model
            dynamic? workbookConnection = null;
            try
            {
                workbookConnection = handle.Workbook.Connections.Add2(
                    Name: $"WorksheetConnection_{tableName}",
                    Description: $"Connection to {tableName}",
                    ConnectionString: $"WORKSHEET;{tableName}",
                    CommandText: tableName,
                    lCmdtype: 2);  // xlCmdTable

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref workbookConnection);
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error adding table to data model: {ex.Message}";
        }
        finally
        {
            ComUtilities.Release(ref table);
        }

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    // === HELPER METHODS ===

    private static dynamic? FindTableColumn(dynamic table, string columnName)
    {
        dynamic? listColumns = null;
        dynamic? column = null;
        try
        {
            listColumns = table.ListColumns;
            for (int i = 1; i <= listColumns.Count; i++)
            {
                dynamic? col = null;
                try
                {
                    col = listColumns.Item(i);
                    if (col.Name == columnName)
                    {
                        column = col;
                        col = null; // Don't release - we're returning it
                        break;
                    }
                }
                finally
                {
                    if (col != null)
                    {
                        ComUtilities.Release(ref col);
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref listColumns);
        }
        return column;
    }

    private static int ParseTotalFunction(string totalFunction)
    {
        return totalFunction.ToLowerInvariant() switch
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
    }
}
