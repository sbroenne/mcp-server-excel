using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult SetPageBreaks(IExcelBatch batch, string sheetName, PageBreakOptions pageBreakOptions)
    {
        ArgumentNullException.ThrowIfNull(pageBreakOptions);
        ValidatePageBreakPositions(pageBreakOptions.Rows, 1048576, "rows");
        ValidatePageBreakPositions(pageBreakOptions.Columns, 16384, "columns");
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.HPageBreaks? horizontal = null;
            Excel.VPageBreaks? vertical = null;
            Excel.Range? cells = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                if (sheet.ProtectContents)
                    throw new InvalidOperationException("Unprotect the worksheet before replacing manual page breaks.");
                horizontal = sheet.HPageBreaks;
                vertical = sheet.VPageBreaks;
                cells = sheet.Cells;
                ct.ThrowIfCancellationRequested();
                sheet.ResetAllPageBreaks();
                foreach (int row in pageBreakOptions.Rows)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? anchor = null;
                    Excel.HPageBreak? pageBreak = null;
                    try
                    {
                        anchor = cells[row, 1];
                        pageBreak = horizontal.Add(anchor);
                    }
                    finally
                    {
                        ComUtilities.Release(ref pageBreak);
                        ComUtilities.Release(ref anchor);
                    }
                }
                foreach (int column in pageBreakOptions.Columns)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? anchor = null;
                    Excel.VPageBreak? pageBreak = null;
                    try
                    {
                        anchor = cells[1, column];
                        pageBreak = vertical.Add(anchor);
                    }
                    finally
                    {
                        ComUtilities.Release(ref pageBreak);
                        ComUtilities.Release(ref anchor);
                    }
                }
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref vertical);
                ComUtilities.Release(ref horizontal);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SheetPageBreaksResult GetPageBreaks(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PageSetup? setup = null;
            Excel.HPageBreaks? horizontal = null;
            Excel.VPageBreaks? vertical = null;
            try
            {
                sheet = FindRequiredSheet(ctx.Book, sheetName);
                setup = sheet.PageSetup;
                var result = new SheetPageBreaksResult
                {
                    FilePath = batch.WorkbookPath,
                    SheetName = sheetName,
                    PrintArea = setup.PrintArea ?? string.Empty
                };
                horizontal = sheet.HPageBreaks;
                vertical = sheet.VPageBreaks;
                for (int index = 1; index <= horizontal.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.HPageBreak? pageBreak = null;
                    Excel.Range? location = null;
                    try
                    {
                        pageBreak = horizontal[index];
                        location = pageBreak.Location;
                        result.Horizontal.Add(new PageBreakInfo
                        {
                            Position = location.Row,
                            Address = location.Address,
                            IsManual = pageBreak.Type == Excel.XlPageBreak.xlPageBreakManual,
                            Extent = pageBreak.Extent.ToString()
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref location);
                        ComUtilities.Release(ref pageBreak);
                    }
                }
                for (int index = 1; index <= vertical.Count; index++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.VPageBreak? pageBreak = null;
                    Excel.Range? location = null;
                    try
                    {
                        pageBreak = vertical[index];
                        location = pageBreak.Location;
                        result.Vertical.Add(new PageBreakInfo
                        {
                            Position = location.Column,
                            Address = location.Address,
                            IsManual = pageBreak.Type == Excel.XlPageBreak.xlPageBreakManual,
                            Extent = pageBreak.Extent.ToString()
                        });
                    }
                    finally
                    {
                        ComUtilities.Release(ref location);
                        ComUtilities.Release(ref pageBreak);
                    }
                }
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref vertical);
                ComUtilities.Release(ref horizontal);
                ComUtilities.Release(ref setup);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void ValidatePageBreakPositions(List<int>? positions, int maximum, string name)
    {
        if (positions is null)
            throw new ArgumentException($"Page break {name} list is required; use an empty list to clear.");
        if (positions.Count > 1026 || positions.Any(position => position < 2 || position > maximum) ||
            positions.Distinct().Count() != positions.Count)
            throw new ArgumentException($"Page break {name} must contain at most 1026 distinct positions from 2 through {maximum}.");
    }
}
