#pragma warning disable IDE0005 // Using directive is unnecessary (all usings are needed for COM interop)

using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table structure operations (Resize, ToggleTotals, SetColumnTotal, SetStyle)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public void Resize(IExcelBatch batch, string tableName, string newRange)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? newRangeObj = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                sheet = table.Parent;
                newRangeObj = sheet.Range[newRange];

                // Resize the table
                table.Resize(newRangeObj);

                return 0;
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
    public void ToggleTotals(IExcelBatch batch, string tableName, bool showTotals)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                table.ShowTotals = showTotals;

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public void SetColumnTotal(IExcelBatch batch, string tableName, string columnName, string totalFunction)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? listColumns = null;
            dynamic? column = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
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
                    throw new InvalidOperationException($"Column '{columnName}' not found in table '{tableName}'");
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

                return 0;
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
    public void SetStyle(IExcelBatch batch, string tableName, string tableStyle)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found");
                }

                table.TableStyle = tableStyle;

                return 0;
            }
            finally
            {
                ComUtilities.Release(ref table);
            }
        });
    }
}

