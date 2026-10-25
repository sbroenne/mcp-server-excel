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
        ValidateRequiredTableName(tableName);

        // Resolve rows from inline parameter or file
        var resolvedRows = ParameterTransforms.ResolveValuesOrFile(rows, rowsFile, "rows");

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? dataBodyRange = null;
            Excel.Range? dataRows = null;
            Excel.ListColumns? listColumns = null;
            Excel.Range? tableRange = null;
            Excel.Range? cells = null;
            int originalCalculation = -1;
            bool calculationChanged = false;
            Exception? operationFailure = null;

            try
            {
                table = FindTable(ctx.Book, tableName);

                sheet = (Excel.Worksheet)(object)table.Parent;

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
                    dataRows = dataBodyRange.Rows;
                    currentRow = dataBodyRange.Row + dataRows.Count;
                }
                else
                {
                    // Table has only headers
                    Excel.Range? headerRange = null;
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

                listColumns = table.ListColumns;
                int columnCount = listColumns.Count;
                for (int rowIndex = 0; rowIndex < resolvedRows.Count; rowIndex++)
                {
                    var row = resolvedRows[rowIndex];
                    if (row is null || row.Count != columnCount)
                    {
                        throw new ArgumentException(
                            $"Append row {rowIndex + 1} must have {columnCount} columns to match table '{tableName}'.",
                            nameof(rows));
                    }
                }
                tableRange = table.Range;
                int tableRow = tableRange.Row;
                int tableColumn = tableRange.Column;
                cells = sheet.Cells;
                int rowsToAdd = resolvedRows.Count;

                // Calculation suppressed here (not in ExcelWriteGuard) because Data Model ops need it enabled
                originalCalculation = (int)ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = (Excel.XlCalculation)(-4135);
                    calculationChanged = true;
                }

                // Write data to cells below the table
                for (int i = 0; i < resolvedRows.Count; i++)
                {
                    var rowValues = resolvedRows[i];
                    for (int j = 0; j < columnCount; j++)
                    {
                        Excel.Range? cell = null;
                        try
                        {
                            cell = (Excel.Range)cells[currentRow + i, tableColumn + j];
                            cell.Value2 = RangeHelpers.ConvertToCellValue(rowValues[j]);
                        }
                        finally
                        {
                            ComUtilities.Release(ref cell);
                        }
                    }
                }

                // Restore calculation before Resize (table operations may need recalculation)
                if (calculationChanged && originalCalculation != -1)
                {
                    ctx.App.Calculation = (Excel.XlCalculation)originalCalculation;
                    calculationChanged = false;
                }

                // Resize table to include new rows
                int newLastRow = currentRow + rowsToAdd - 1;
                int newLastCol = tableColumn + columnCount - 1;
                Excel.Range? firstCell = null;
                Excel.Range? lastCell = null;
                string newRangeAddress;
                try
                {
                    firstCell = (Excel.Range)cells[tableRow, tableColumn];
                    lastCell = (Excel.Range)cells[newLastRow, newLastCol];
                    newRangeAddress = $"{firstCell.Address}:{lastCell.Address}";
                }
                finally
                {
                    ComUtilities.Release(ref lastCell);
                    ComUtilities.Release(ref firstCell);
                }

                Excel.Range? resizeRange = null;
                try
                {
                    resizeRange = sheet.Range[newRangeAddress];
                    table.Resize(resizeRange);
                }
                finally
                {
                    ComUtilities.Release(ref resizeRange);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                try
                {
                    if (calculationChanged && originalCalculation != -1)
                    {
                        OperationCleanup.Run(() => ctx.App.Calculation = (Excel.XlCalculation)originalCalculation,
                            operationFailure);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref cells);
                    ComUtilities.Release(ref tableRange);
                    ComUtilities.Release(ref listColumns);
                    ComUtilities.Release(ref dataRows);
                    ComUtilities.Release(ref dataBodyRange);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref table);
                }
            }
        });
    }

    /// <inheritdoc />
    public TableDataResult GetData(IExcelBatch batch, string tableName, bool visibleOnly)
    {
        ValidateRequiredTableName(tableName);

        var result = new TableDataResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumns? listColumns = null;
            Excel.ListRows? listRows = null;
            Excel.Range? dataBodyRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                listColumns = table.ListColumns;
                int columnCount = listColumns.Count;
                result.ColumnCount = columnCount;

                for (int i = 1; i <= columnCount; i++)
                {
                    Excel.ListColumn? column = null;
                    try
                    {
                        column = listColumns[i];
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

    private static bool IsListRowVisible(Excel.ListRows? listRows, int listRowCount, int index)
    {
        if (listRows is null || index > listRowCount)
        {
            return true;
        }

        Excel.ListRow? listRow = null;
        Excel.Range? rowRange = null;
        Excel.Range? entireRow = null;
        try
        {
            listRow = listRows[index];
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
