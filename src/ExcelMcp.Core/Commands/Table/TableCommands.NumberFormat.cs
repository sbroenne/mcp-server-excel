using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// TableCommands partial class - Number formatting operations
/// Delegates to RangeCommands for actual formatting operations
/// </summary>
public partial class TableCommands
{
    private readonly IRangeCommands _rangeCommands = new RangeCommands();

    // === NUMBER FORMATTING OPERATIONS ===

    /// <inheritdoc />
    public async Task<RangeNumberFormatResult> GetColumnNumberFormatAsync(IExcelBatch batch, string tableName, string columnName)
    {
        // First, get the table's sheet name and column range
        var columnRange = await GetColumnRangeAsync(batch, tableName, columnName);
        
        if (!columnRange.Success)
        {
            return new RangeNumberFormatResult
            {
                Success = false,
                ErrorMessage = columnRange.ErrorMessage,
                FilePath = batch.WorkbookPath
            };
        }

        // Delegate to RangeCommands to get number formats
        return await _rangeCommands.GetNumberFormatsAsync(batch, columnRange.SheetName, columnRange.RangeAddress);
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetColumnNumberFormatAsync(IExcelBatch batch, string tableName, string columnName, string formatCode)
    {
        // First, get the table's sheet name and column data range (excludes header)
        var columnRange = await GetColumnDataRangeAsync(batch, tableName, columnName);
        
        if (!columnRange.Success)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = columnRange.ErrorMessage,
                FilePath = batch.WorkbookPath,
                Action = "set-column-number-format"
            };
        }

        // Delegate to RangeCommands to set number format
        var result = await _rangeCommands.SetNumberFormatAsync(batch, columnRange.SheetName, columnRange.RangeAddress, formatCode);
        
        result.Action = "set-column-number-format";
        result.SuggestedNextActions =
        [
            "Use 'info' to verify table structure",
            $"Use 'get-column-number-format' to verify format applied to {columnName}",
            "Use range 'get-values' to see formatted values"
        ];
        result.WorkflowHint = $"Applied format '{formatCode}' to table '{tableName}' column '{columnName}'";

        return result;
    }

    // === HELPER METHODS ===

    /// <summary>
    /// Gets the full column range (including header) for a table column
    /// </summary>
    private async Task<TableColumnRangeResult> GetColumnRangeAsync(IExcelBatch batch, string tableName, string columnName)
    {
        var result = new TableColumnRangeResult
        {
            FilePath = batch.WorkbookPath
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? column = null;
            dynamic? columnRange = null;
            try
            {
                table = FindTable(ctx.Book, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' not found";
                    return result;
                }

                // Find the column
                column = FindColumn(table, columnName);
                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{columnName}' not found in table '{tableName}'";
                    return result;
                }

                // Get the entire column range (including header)
                columnRange = column.Range;
                
                result.SheetName = columnRange.Worksheet.Name;
                result.RangeAddress = columnRange.Address;
                result.Success = true;

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to get column range: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref columnRange);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Gets the data range (excluding header) for a table column
    /// </summary>
    private async Task<TableColumnRangeResult> GetColumnDataRangeAsync(IExcelBatch batch, string tableName, string columnName)
    {
        var result = new TableColumnRangeResult
        {
            FilePath = batch.WorkbookPath
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? table = null;
            dynamic? column = null;
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

                // Find the column
                column = FindColumn(table, columnName);
                if (column == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{columnName}' not found in table '{tableName}'";
                    return result;
                }

                // Get the data body range (excludes header and totals)
                dataBodyRange = column.DataBodyRange;
                
                if (dataBodyRange == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Table '{tableName}' has no data rows";
                    return result;
                }

                result.SheetName = dataBodyRange.Worksheet.Name;
                result.RangeAddress = dataBodyRange.Address;
                result.Success = true;

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Failed to get column data range: {ex.Message}";
                return result;
            }
            finally
            {
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Finds a column in a table by name
    /// </summary>
    private static dynamic? FindColumn(dynamic table, string columnName)
    {
        dynamic? columns = null;
        try
        {
            columns = table.ListColumns;
            int count = Convert.ToInt32(columns.Count);

            for (int i = 1; i <= count; i++)
            {
                dynamic? col = null;
                try
                {
                    col = columns.Item(i);
                    string name = col.Name?.ToString() ?? "";
                    
                    if (name.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Return without releasing - caller will release
                        var result = col;
                        col = null; // Prevent release in finally
                        return result;
                    }
                }
                finally
                {
                    if (col != null) ComUtilities.Release(ref col);
                }
            }

            return null;
        }
        finally
        {
            ComUtilities.Release(ref columns);
        }
    }
}

/// <summary>
/// Helper result for internal table column range resolution
/// </summary>
internal class TableColumnRangeResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
}
