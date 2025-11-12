using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// FilePath-based methods for Excel Table (ListObject) management - uses FileHandleManager
/// </summary>
public partial class TableCommands
{
    #region Lifecycle Operations

    /// <summary>
    /// Lists all Excel Tables in the workbook (filePath-based)
    /// </summary>
    public async Task<TableListResult> ListAsync(string filePath)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        return await handle.ExecuteAsync((app, book) =>
        {
            var result = new TableListResult();
            dynamic? sheets = null;

            try
            {
                sheets = book.Worksheets;
                var tables = new List<TableInfo>();

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
                                tables.Add(new TableInfo
                                {
                                    Name = table.Name,
                                    SheetName = sheet.Name,
                                    Range = table.Range.Address
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

                result.Tables = tables;
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to list tables: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }

            return ValueTask.FromResult(result);
        });
    }

    /// <summary>
    /// Creates a new Excel Table from a range (filePath-based)
    /// </summary>
    public async Task<OperationResult> CreateAsync(string filePath, string sheetName, string tableName, string range, bool hasHeaders = true, string? tableStyle = null)
    {
        ValidateTableName(tableName);

        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? listObjects = null;
            dynamic? sourceRange = null;
            dynamic? newTable = null;

            try
            {
                sheets = book.Worksheets;
                sheet = sheets.Item(sheetName);
                listObjects = sheet.ListObjects;

                // Check if table name already exists
                var existingTable = FindTable(book, tableName);
                if (existingTable != null)
                {
                    ComUtilities.Release(ref existingTable);
                    throw new InvalidOperationException($"Table '{tableName}' already exists");
                }

                sourceRange = sheet.Range[range];
                newTable = listObjects.Add(
                    XlListObjectSourceType: 1, // xlSrcRange
                    Source: sourceRange,
                    XlListObjectHasHeaders: hasHeaders ? 1 : 2
                );

                newTable.Name = tableName;

                if (!string.IsNullOrEmpty(tableStyle))
                {
                    newTable.TableStyle = tableStyle;
                }

                opResult.Success = true;
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to create table: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref newTable);
                ComUtilities.Release(ref sourceRange);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }

            return ValueTask.FromResult(opResult);
        });

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
        ValidateTableName(newName);

        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? table = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                // Check if new name already exists
                var existingTable = FindTable(book, newName);
                if (existingTable != null)
                {
                    ComUtilities.Release(ref existingTable);
                    throw new InvalidOperationException($"Table '{newName}' already exists");
                }

                table.Name = newName;
                opResult.Success = true;
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to rename table: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(opResult);
        });

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Deletes an Excel Table (converts back to range) (filePath-based)
    /// </summary>
    public async Task<OperationResult> DeleteAsync(string filePath, string tableName)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? table = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                table.Unlist();
                opResult.Success = true;
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to delete table: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(opResult);
        });

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
        return await handle.ExecuteAsync((app, book) =>
        {
            var result = new TableInfoResult();
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                sheet = table.Parent;
                range = table.Range;

                result.Name = table.Name;
                result.SheetName = sheet.Name;
                result.Range = range.Address;
                result.HasHeaders = table.ShowHeaders;
                result.HasTotals = table.ShowTotals;
                result.TableStyle = ComUtilities.SafeGetString(table, "TableStyle");
                result.Success = true;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to get table info: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(result);
        });
    }

    /// <summary>
    /// Resizes an Excel Table to a new range (filePath-based)
    /// </summary>
    public async Task<OperationResult> ResizeAsync(string filePath, string tableName, string newRange)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? range = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                sheet = table.Parent;
                range = sheet.Range[newRange];
                table.Resize(range);
                opResult.Success = true;
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to resize table: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(opResult);
        });

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    #endregion

    #region Styling and Totals

    /// <summary>
    /// Toggles the totals row for an Excel Table (filePath-based)
    /// </summary>
    public async Task<OperationResult> ToggleTotalsAsync(string filePath, string tableName, bool showTotals)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? table = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                table.ShowTotals = showTotals;
                opResult.Success = true;
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to toggle totals: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(opResult);
        });

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    /// <summary>
    /// Sets the totals function for a specific column in an Excel Table (filePath-based)
    /// </summary>
    public async Task<OperationResult> SetColumnTotalAsync(string filePath, string tableName, string columnName, string totalFunction)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? table = null;
            dynamic? column = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                if (!table.ShowTotals)
                {
                    throw new InvalidOperationException($"Table '{tableName}' does not have totals row enabled");
                }

                column = table.ListColumns.Item(columnName);
                int totalsCalculation = totalFunction.ToLowerInvariant() switch
                {
                    "sum" => 1,
                    "average" or "avg" => 2,
                    "count" => 3,
                    "countnums" => 4,
                    "min" => 5,
                    "max" => 6,
                    "stddev" => 7,
                    "var" => 8,
                    "none" => 0,
                    _ => throw new ArgumentException($"Invalid total function: {totalFunction}")
                };

                column.TotalsCalculation = totalsCalculation;
                opResult.Success = true;
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to set column total: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(opResult);
        });

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
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? table = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                table.TableStyle = tableStyle;
                opResult.Success = true;
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to set table style: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(opResult);
        });

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    #endregion

    #region Data Operations

    /// <summary>
    /// Appends rows to an Excel Table (table auto-expands) (filePath-based)
    /// </summary>
    public async Task<OperationResult> AppendAsync(string filePath, string tableName, List<List<object?>> rows)
    {
        var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? table = null;
            dynamic? dataBodyRange = null;
            dynamic? firstRowAfter = null;
            dynamic? appendRange = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                dataBodyRange = table.DataBodyRange;
                int lastRow = dataBodyRange.Row + dataBodyRange.Rows.Count - 1;
                int firstCol = dataBodyRange.Column;
                int columnCount = dataBodyRange.Columns.Count;

                if (rows[0].Count != columnCount)
                {
                    throw new InvalidOperationException($"Row has {rows[0].Count} columns but table has {columnCount} columns");
                }

                // Convert to 2D array
                object[,] data = new object[rows.Count, columnCount];
                for (int i = 0; i < rows.Count; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        data[i, j] = rows[i][j] ?? "";
                    }
                }

                dynamic sheet = table.Parent;
                firstRowAfter = sheet.Cells[lastRow + 1, firstCol];
                appendRange = firstRowAfter.Resize[rows.Count, columnCount];
                appendRange.Value2 = data;

                ComUtilities.Release(ref sheet);
                opResult.Success = true;
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to append rows: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref appendRange);
                ComUtilities.Release(ref firstRowAfter);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(opResult);
        });

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
        var result = await handle.ExecuteAsync((app, book) =>
        {
            var opResult = new OperationResult();
            dynamic? table = null;

            try
            {
                table = FindTable(book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                // Add table to workbook connections
                string connString = $"WORKSHEET;{table.Parent.Name}";
                dynamic? connections = book.Connections;
                dynamic? workbookConn = null;

                try
                {
                    workbookConn = connections.Add2(
                        Name: $"WorksheetConnection_{tableName}",
                        Description: $"Connection to table {tableName}",
                        ConnectionString: connString,
                        CommandText: tableName,
                        lCmdtype: 2 // xlCmdTable
                    );

                    workbookConn.OLEDBConnection.BackgroundQuery = false;

                    opResult.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref workbookConn);
                    ComUtilities.Release(ref connections);
                }
            }
            catch (Exception ex)
            {
                opResult.Success = false;
                opResult.ErrorMessage = $"Failed to add table to data model: {ex.Message}";
            }
            finally
            {
                ComUtilities.Release(ref table);
            }

            return ValueTask.FromResult(opResult);
        });

        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(filePath);
        }

        return result;
    }

    #endregion

    // NOTE: Filter operations, column operations, structured references, sort operations, 
    // and number formatting methods follow the same pattern but are omitted for brevity.
    // They will be added in the actual implementation.
}
