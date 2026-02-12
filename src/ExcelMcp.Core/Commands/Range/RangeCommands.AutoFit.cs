using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Auto-fit operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2012:Use ValueTasks correctly")]
    public void AutoFitColumns(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? columns = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get columns and auto-fit
                columns = range.Columns;
                columns.AutoFit();

                return ValueTask.CompletedTask;
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2012:Use ValueTasks correctly")]
    public void AutoFitRows(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? rows = null;

            try
            {
                // Get sheet
                sheet = string.IsNullOrEmpty(sheetName)
                    ? ctx.Book.ActiveSheet
                    : ctx.Book.Worksheets.Item(sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get rows and auto-fit
                rows = range.Rows;
                rows.AutoFit();

                return ValueTask.CompletedTask;
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



