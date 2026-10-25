using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// TableCommands partial class - Number formatting operations
/// Delegates to RangeCommands for actual formatting operations
/// </summary>
public partial class TableCommands
{
    private readonly RangeCommands _rangeCommands = new RangeCommands();

    // === NUMBER FORMATTING OPERATIONS ===

    /// <inheritdoc />
    public RangeNumberFormatResult GetColumnNumberFormat(IExcelBatch batch, string tableName, string columnName)
    {
        // First, get the table's sheet name and column range
        var columnRange = GetColumnRange(batch, tableName, columnName);

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
        return _rangeCommands.GetNumberFormats(batch, columnRange.SheetName, columnRange.RangeAddress);
    }

    /// <inheritdoc />
    public OperationResult SetColumnNumberFormat(IExcelBatch batch, string tableName, string columnName, string formatCode)
    {
        // First, get the table's sheet name and column data range (excludes header)
        var columnRange = GetColumnDataRange(batch, tableName, columnName);

        if (!columnRange.Success)
        {
            throw new InvalidOperationException(columnRange.ErrorMessage ?? "Failed to get column range");
        }

        // Delegate to RangeCommands to set number format
        return _rangeCommands.SetNumberFormat(batch, columnRange.SheetName, columnRange.RangeAddress, formatCode);
    }

    // === HELPER METHODS ===

    /// <summary>
    /// Gets the full column range (including header) for a table column
    /// </summary>
    private static TableColumnRangeResult GetColumnRange(IExcelBatch batch, string tableName, string columnName)
    {
        var result = new TableColumnRangeResult
        {
            FilePath = batch.WorkbookPath
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumn? column = null;
            Excel.Range? columnRange = null;
            Excel.Worksheet? sheet = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Find the column
                column = FindColumn(table, columnName);
                if (column == null)
                {
                    throw new InvalidOperationException($"Column '{columnName}' not found in table '{tableName}'");
                }

                // Get the entire column range (including header)
                columnRange = column.Range;

                sheet = columnRange.Worksheet;
                result.SheetName = sheet.Name;
                result.RangeAddress = columnRange.Address;
                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref columnRange);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Gets the data range (excluding header) for a table column
    /// </summary>
    private static TableColumnRangeResult GetColumnDataRange(IExcelBatch batch, string tableName, string columnName)
    {
        var result = new TableColumnRangeResult
        {
            FilePath = batch.WorkbookPath
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.ListColumn? column = null;
            Excel.Range? dataBodyRange = null;
            Excel.Worksheet? sheet = null;
            try
            {
                table = FindTable(ctx.Book, tableName);

                // Find the column
                column = FindColumn(table, columnName);
                if (column == null)
                {
                    throw new InvalidOperationException($"Column '{columnName}' not found in table '{tableName}'");
                }

                // Get the data body range (excludes header and totals)
                dataBodyRange = column.DataBodyRange;

                if (dataBodyRange == null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' has no data rows");
                }

                sheet = dataBodyRange.Worksheet;
                result.SheetName = sheet.Name;
                result.RangeAddress = dataBodyRange.Address;
                result.Success = true;

                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref dataBodyRange);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref table);
            }
        });
    }

    /// <summary>
    /// Finds a column in a table by name
    /// </summary>
    private static Excel.ListColumn? FindColumn(Excel.ListObject table, string columnName)
    {
        Excel.ListColumns? columns = null;
        try
        {
            columns = table.ListColumns;
            int count = Convert.ToInt32(columns.Count);

            for (int i = 1; i <= count; i++)
            {
                Excel.ListColumn? col = null;
                try
                {
                    col = columns[i];
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
                    ComUtilities.Release(ref col);
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
internal sealed class TableColumnRangeResult : ResultBase
{
    public string SheetName { get; set; } = string.Empty;
    public string RangeAddress { get; set; } = string.Empty;
}

