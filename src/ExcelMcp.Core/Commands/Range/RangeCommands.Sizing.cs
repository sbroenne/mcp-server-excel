using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Explicit sizing operations for Excel ranges (partial class)
/// Sets specific column widths and row heights.
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult SetColumnWidth(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        double columnWidth)
    {
        if (columnWidth < 0.25 || columnWidth > 409)
        {
            throw new ArgumentException("columnWidth must be between 0.25 and 409 points", nameof(columnWidth));
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? columns = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                // Get range and its columns
                range = sheet.Range[rangeAddress];
                columns = range.Columns;

                // Set column width for all columns in range
                columns.ColumnWidth = columnWidth;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref columns!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetRowHeight(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        double rowHeight)
    {
        if (rowHeight < 0 || rowHeight > 409)
        {
            throw new ArgumentException("rowHeight must be between 0 and 409 points", nameof(rowHeight));
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? rows = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets[sheetName];

                // Get range and its rows
                range = sheet.Range[rangeAddress];
                rows = range.Rows;

                // Set row height for all rows in range
                rows.RowHeight = rowHeight;

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref rows!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }
}
