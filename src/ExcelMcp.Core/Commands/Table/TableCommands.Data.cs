using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table data operations (AppendRows)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public OperationResult Append(IExcelBatch batch, string tableName, List<List<object?>>? rows = null, string? rowsFile = null)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        // Resolve rows from inline parameter or file
        var resolvedRows = ParameterTransforms.ResolveValuesOrFile(rows, rowsFile, "rows");

        return batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? dataBodyRange = null;
            dynamic? tableRange = null;
            dynamic? writeRange = null;
            dynamic? resizeRange = null;
            dynamic? startCell = null;
            dynamic? endCell = null;
            dynamic? rollbackRange = null;
            int originalCalculation = -1;
            bool calculationChanged = false;
            bool resized = false;
            bool appendSucceeded = false;
            string? originalTableAddress = null;

            try
            {
                table = FindTable(ctx.Book, tableName);

                sheet = table.Parent;

                // Validate data
                if (resolvedRows.Count == 0)
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

                tableRange = table.Range;
                int tableRow = tableRange.Row;
                int tableColumn = tableRange.Column;
                originalTableAddress = Convert.ToString(
                    tableRange.Address,
                    CultureInfo.InvariantCulture);
                int columnCount = table.ListColumns.Count;
                int rowsToAdd = resolvedRows.Count;

                for (int rowIndex = 0; rowIndex < rowsToAdd; rowIndex++)
                {
                    int actualColumnCount = resolvedRows[rowIndex].Count;
                    if (actualColumnCount != columnCount)
                    {
                        throw new ArgumentException(
                            $"Row {rowIndex + 1} column count ({actualColumnCount}) doesn't match table column count ({columnCount})",
                            nameof(rows));
                    }
                }

                // Excel COM accepts a whole rectangular block in one call. Build the
                // 1-based SAFEARRAY before mutating the workbook so conversion errors
                // cannot leave a partially appended table.
                object[,] arrayValues = (object[,])Array.CreateInstance(
                    typeof(object),
                    [rowsToAdd, columnCount],
                    [1, 1]);
                for (int row = 1; row <= rowsToAdd; row++)
                {
                    for (int column = 1; column <= columnCount; column++)
                    {
                        arrayValues[row, column] = RangeHelpers.ConvertToCellValue(
                            resolvedRows[row - 1][column - 1]);
                    }
                }

                // Calculation suppressed here (not in ExcelWriteGuard) because Data Model ops need it enabled
                originalCalculation = (int)ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = (Excel.XlCalculation)(-4135);
                    calculationChanged = true;
                }

                // Resize once, then write the new rows as one contiguous block. If a
                // totals row is visible Excel moves it below the newly appended rows.
                int newLastRow = currentRow + rowsToAdd - 1;
                int newLastCol = tableColumn + columnCount - 1;
                bool showTotals = table.ShowTotals;

                startCell = sheet.Cells[tableRow, tableColumn];
                endCell = sheet.Cells[newLastRow + (showTotals ? 1 : 0), newLastCol];
                string resizeAddress = $"{startCell.Address}:{endCell.Address}";
                resizeRange = sheet.Range[resizeAddress];
                table.Resize(resizeRange);
                resized = true;

                ComUtilities.Release(ref startCell);
                ComUtilities.Release(ref endCell);

                startCell = sheet.Cells[currentRow, tableColumn];
                endCell = sheet.Cells[newLastRow, newLastCol];
                string writeAddress = $"{startCell.Address}:{endCell.Address}";
                writeRange = sheet.Range[writeAddress];
                writeRange.Value2 = arrayValues;
                appendSucceeded = true;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                if (!appendSucceeded && resized && !string.IsNullOrWhiteSpace(originalTableAddress))
                {
                    try
                    {
                        rollbackRange = sheet.Range[originalTableAddress];
                        table.Resize(rollbackRange);
                    }
                    catch (Exception)
                    {
                        // Rollback is best effort. Preserve the original append failure
                        // while ensuring the temporary range is released below.
                    }
                    finally
                    {
                        ComUtilities.Release(ref rollbackRange);
                    }
                }

                if (calculationChanged && originalCalculation != -1)
                {
                    try
                    {
                        ctx.App.Calculation = (Excel.XlCalculation)originalCalculation;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Ignore errors restoring calculation mode
                    }
                }
                ComUtilities.Release(ref endCell);
                ComUtilities.Release(ref startCell);
                ComUtilities.Release(ref resizeRange);
                ComUtilities.Release(ref writeRange);
                ComUtilities.Release(ref tableRange);
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
                        result.Data.Add([rawValues]);
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



