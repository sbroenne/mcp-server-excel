using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Auto-fit operations for Excel ranges (partial class)
/// </summary>
public partial class RangeCommands
{
    /// <inheritdoc />
    public OperationResult AutoFitColumns(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Range? columns = null;

            try
            {
                // Get sheet
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get columns and auto-fit
                columns = range.Columns;
                columns.AutoFit();

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
    public OperationResult AutoFitRows(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Range? rows = null;

            try
            {
                // Get sheet
                sheet = Utilities.WorksheetLookup.Get(ctx.Book, sheetName);

                // Get range
                range = sheet.Range[rangeAddress];

                // Get rows and auto-fit
                rows = range.Rows;
                rows.AutoFit();

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

