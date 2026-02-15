using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;

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
            int originalCalculation = -1; // xlCalculationAutomatic = -4105, xlCalculationManual = -4135
            bool calculationChanged = false;

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

                int columnCount = table.ListColumns.Count;
                int rowsToAdd = resolvedRows.Count;

                // CRITICAL: Temporarily disable automatic calculation to prevent Excel from
                // hanging when appended data triggers dependent formulas that reference Data Model/DAX.
                // Without this, setting values can block the COM interface during recalculation.
                originalCalculation = ctx.App.Calculation;
                if (originalCalculation != -4135) // xlCalculationManual
                {
                    ctx.App.Calculation = -4135; // xlCalculationManual
                    calculationChanged = true;
                }

                // Write data to cells below the table
                for (int i = 0; i < resolvedRows.Count; i++)
                {
                    var rowValues = resolvedRows[i];
                    for (int j = 0; j < Math.Min(rowValues.Count, columnCount); j++)
                    {
                        dynamic? cell = null;
                        try
                        {
                            cell = sheet.Cells[currentRow + i, table.Range.Column + j];
                            cell.Value2 = rowValues[j] ?? string.Empty;
                        }
                        finally
                        {
                            ComUtilities.Release(ref cell);
                        }
                    }
                }

                // Restore calculation before resize so the table can recalculate after the operation
                if (calculationChanged && originalCalculation != -1)
                {
                    try
                    {
                        ctx.App.Calculation = originalCalculation;
                        calculationChanged = false; // Mark as restored
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Ignore errors restoring calculation mode - will try again in finally
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

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                // Restore original calculation mode if not already restored
                if (calculationChanged && originalCalculation != -1)
                {
                    try
                    {
                        ctx.App.Calculation = originalCalculation;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Ignore errors restoring calculation mode - not critical
                    }
                }
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



