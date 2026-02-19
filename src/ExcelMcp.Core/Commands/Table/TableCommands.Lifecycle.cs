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
    public TableListResult List(IExcelBatch batch)
    {
        var result = new TableListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
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
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Create(IExcelBatch batch, string sheetName, string tableName, string rangeAddress, bool hasHeaders = true, string? tableStyle = null)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        return batch.Execute((ctx, ct) =>
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
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                // Check if table name already exists
                if (TableExists(ctx.Book, tableName))
                {
                    throw new InvalidOperationException($"Table '{tableName}' already exists");
                }

                // Get the range to convert to table
                rangeObj = sheet.Range[rangeAddress];

                // Auto-expand single cell to current region (common UX pattern)
                // This allows users to specify just "A1" instead of the full range
                dynamic? currentRegion = null;
                try
                {
                    // Check if single cell (no colon in address = single cell)
                    if (!rangeAddress.Contains(':'))
                    {
                        currentRegion = rangeObj.CurrentRegion;
                        if (currentRegion != null && currentRegion.Cells.Count > 1)
                        {
                            // Use the expanded current region instead
                            ComUtilities.Release(ref rangeObj);
                            rangeObj = currentRegion;
                            currentRegion = null; // Don't release twice
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref currentRegion);
                }

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

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
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
    public OperationResult Rename(IExcelBatch batch, string tableName, string newName)
    {
        // Security: Validate table names
        ValidateTableName(tableName);
        ValidateTableName(newName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Check if new name already exists
                if (TableExists(ctx.Book, newName))
                {
                    throw new InvalidOperationException($"Table '{newName}' already exists");
                }

                table.Name = newName;
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult Delete(IExcelBatch batch, string tableName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? tableRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // SECURITY FIX: Store range info before Unlist() for proper cleanup
                try
                {
                    tableRange = table.Range;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Ignore if range is not accessible
                }

                // Convert table back to range (Unlist)
                table.Unlist();

                // SECURITY FIX: After Unlist(), we must explicitly release the table COM object
                // The table object is no longer valid but still holds a COM reference
                ComUtilities.Release(ref table);

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public TableInfoResult Read(IExcelBatch batch, string tableName)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new TableInfoResult { FilePath = batch.WorkbookPath };
        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? dataBodyRange = null;
            dynamic? headerRowRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

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
}



