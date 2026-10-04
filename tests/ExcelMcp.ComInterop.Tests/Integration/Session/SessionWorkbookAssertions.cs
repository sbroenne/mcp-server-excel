using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

internal static class SessionWorkbookAssertions
{
    internal static void AssertIdentity(IExcelBatch batch, string path)
    {
        Assert.Equal(path, batch.WorkbookPath);
        Assert.Equal(path, batch.Execute((context, _) => context.Book.FullName));
        Assert.NotNull(batch.ExcelProcessId);
        Assert.True(batch.IsExcelProcessAlive());
    }

    internal static void WriteMarker(IExcelBatch batch, string marker)
    {
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                cell = sheet.Range["A1"];
                cell.Value2 = marker;
                Assert.Equal(marker, Assert.IsType<string>(cell.Value2));
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    internal static object? ReadMarker(IExcelBatch batch) =>
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                cell = sheet.Range["A1"];
                return cell.Value2;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
