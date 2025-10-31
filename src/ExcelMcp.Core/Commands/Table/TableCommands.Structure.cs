using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table structure operations (Resize, ToggleTotals, SetColumnTotal, SetStyle)
/// </summary>
public partial class TableCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> ResizeAsync(IExcelBatch batch, string tableName, string newRange)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "resize" };
        return await batch.Execute((ctx, ct) =>
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
                result.SuggestedNextActions.Add($"Use 'table-info {tableName}' to verify the new size");
                result.SuggestedNextActions.Add($"Use 'range-get-values' on the table range to view updated data");
                result.WorkflowHint = $"Table '{tableName}' resized to {newRange}";

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
                ComUtilities.Release(ref newRangeObj);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> ToggleTotalsAsync(IExcelBatch batch, string tableName, bool showTotals)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "toggle-totals" };
        return await batch.Execute((ctx, ct) =>
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
    public async Task<OperationResult> SetColumnTotalAsync(IExcelBatch batch, string tableName, string columnName, string totalFunction)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-column-total" };
        return await batch.Execute((ctx, ct) =>
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
                result.SuggestedNextActions.Add($"Use 'table-info {tableName}' to verify totals configuration");
                result.SuggestedNextActions.Add($"Use 'range-get-values' on the table range to see calculated totals");
                result.WorkflowHint = $"Column '{columnName}' total set to {totalFunction}.";

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
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref listColumns);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetStyleAsync(IExcelBatch batch, string tableName, string tableStyle)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "set-style" };
        return await batch.Execute((ctx, ct) =>
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
                result.SuggestedNextActions.Add($"Use 'table-info {tableName}' to verify the style change");
                result.SuggestedNextActions.Add("Common styles: TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11");
                result.WorkflowHint = $"Table '{tableName}' style changed to '{tableStyle}'.";

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
}
