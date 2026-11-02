using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

public partial class RangeCommands
{
    // Numeric values can equal error HRESULTs; confirm the cell's native type before mapping.
    private static bool IsNativeError(ExcelContext context, Excel.Range range, int row, int column)
    {
        Excel.Range? cells = null;
        Excel.Range? cell = null;
        Excel.WorksheetFunction? functions = null;
        try
        {
            cells = range.Cells;
            cell = cells[row, column];
            functions = context.App.WorksheetFunction;
            return functions.IsError(cell);
        }
        finally
        {
            ComUtilities.Release(ref functions);
            ComUtilities.Release(ref cell);
            ComUtilities.Release(ref cells);
        }
    }
}
