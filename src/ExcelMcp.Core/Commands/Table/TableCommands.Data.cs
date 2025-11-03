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
    public async Task<OperationResult> AppendAsync(IExcelBatch batch, string tableName, List<List<object?>> rows)
    {
        // Security: Validate table name
        ValidateTableName(tableName);

        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = "append-rows" };
        return await batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? sheet = null;
            dynamic? dataBodyRange = null;
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

                // Validate data
                if (rows == null || rows.Count == 0)
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
                int rowsToAdd = rows.Count;

                // Write data to cells below the table
                for (int i = 0; i < rows.Count; i++)
                {
                    var rowValues = rows[i];
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
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref table);
            }
        });
    }
}
