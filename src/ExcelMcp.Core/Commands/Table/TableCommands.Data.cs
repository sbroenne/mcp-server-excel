#pragma warning disable IDE0005 // Using directive is unnecessary (all usings are needed for COM interop)

using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table data operations (AppendRows)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public void Append(IExcelBatch batch, string tableName, List<List<object?>> rows)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? dataBodyRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                sheet = table.Parent;

                // Validate data
                if (rows == null || rows.Count == 0)
                {
                    throw new ArgumentException("No data to append", nameof(rows));
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
                int rowsToAdd = rows.Count;

                // PERFORMANCE: Convert rows to 2D array for bulk write (single COM call)
                // This is 100x+ faster than cell-by-cell writing for large datasets
                object[,] arrayValues = (object[,])Array.CreateInstance(typeof(object), [rowsToAdd, columnCount], [1, 1]);

                for (int i = 0; i < rowsToAdd; i++)
                {
                    var rowValues = rows[i];
                    for (int j = 0; j < columnCount; j++)
                    {
                        arrayValues[i + 1, j + 1] = j < rowValues.Count ? (rowValues[j] ?? string.Empty) : string.Empty;
                    }
                }

                // Bulk write to range below the table (single COM call instead of N×M calls)
                dynamic? startCell = null;
                dynamic? endCell = null;
                dynamic? targetRange = null;
                try
                {
                    int startCol = Convert.ToInt32(table.Range.Column);
                    startCell = sheet.Cells[currentRow, startCol];
                    endCell = sheet.Cells[currentRow + rowsToAdd - 1, startCol + columnCount - 1];
                    targetRange = sheet.Range[startCell, endCell];
                    targetRange.Value2 = arrayValues;
                }
                finally
                {
                    ComUtilities.Release(ref targetRange);
                    ComUtilities.Release(ref endCell);
                    ComUtilities.Release(ref startCell);
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

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public TableDataResult GetData(IExcelBatch batch, string tableName, bool visibleOnly)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new TableDataResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? listRows = null;
            dynamic? dataBodyRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                listColumns = table.ListColumns;
                int columnCount = listColumns.Count;
                result.ColumnCount = columnCount;

                for (int i = 1; i <= columnCount; i++)
                {
                    dynamic? column = null;
                    try
                    {
                        column = listColumns.Item(i);
                        string? columnName = column.Name;
                        if (!string.IsNullOrEmpty(columnName))
                        {
                            result.Headers.Add(columnName);
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref column);
                    }
                }

                dataBodyRange = table.DataBodyRange;
                if (dataBodyRange == null)
                {
                    result.Success = true;
                    result.RowCount = 0;
                    return result;
                }

                object? rawValues = dataBodyRange.Value2;
                if (rawValues == null)
                {
                    result.Success = true;
                    result.RowCount = 0;
                    return result;
                }

                listRows = table.ListRows;
                int listRowCount = listRows?.Count ?? 0;

                if (rawValues is object[,] array2D)
                {
                    int rows = array2D.GetLength(0);
                    int cols = array2D.GetLength(1);

                    for (int r = 1; r <= rows; r++)
                    {
                        bool includeRow = !visibleOnly;
                        if (!includeRow)
                        {
                            includeRow = IsListRowVisible(listRows, listRowCount, r);
                        }

                        if (!includeRow)
                        {
                            continue;
                        }

                        var row = new List<object?>(cols);
                        for (int c = 1; c <= cols; c++)
                        {
                            row.Add(array2D[r, c]);
                        }
                        result.Data.Add(row);
                    }
                }
                else
                {
                    bool includeRow = !visibleOnly || IsListRowVisible(listRows, listRowCount, 1);
                    if (includeRow)
                    {
                        result.Data.Add(new List<object?> { rawValues });
                    }
                }

                result.RowCount = result.Data.Count;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref listRows);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    private static bool IsListRowVisible(dynamic? listRows, int listRowCount, int index)
    {
        if (listRows == null || index > listRowCount)
        {
            return true;
        }

        dynamic? listRow = null;
        dynamic? rowRange = null;
        dynamic? entireRow = null;
        try
        {
            listRow = listRows.Item(index);
            rowRange = listRow.Range;
            entireRow = rowRange.EntireRow;

            object? hiddenValue = entireRow.Hidden;
            bool hidden = hiddenValue switch
            {
                bool b => b,
                null => false,
                string s when bool.TryParse(s, out var parsed) => parsed,
                IConvertible convertible => convertible.ToBoolean(CultureInfo.InvariantCulture),
                _ => false
            };

            return !hidden;
        }
        finally
        {
            ComUtilities.Release(ref entireRow);
            ComUtilities.Release(ref rowRange);
            ComUtilities.Release(ref listRow);
        }
    }
}

