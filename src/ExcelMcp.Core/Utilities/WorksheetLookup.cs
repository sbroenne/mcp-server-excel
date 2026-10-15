using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class WorksheetLookup
{
    internal static Excel.Worksheet Get(Excel.Workbook workbook, string? name)
    {
        if (string.IsNullOrEmpty(name))
            return (Excel.Worksheet)workbook.ActiveSheet;

        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
            return (Excel.Worksheet)sheets[name];
        }
        finally
        {
            ComUtilities.Release(ref sheets);
        }
    }
}
