using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

internal static class PersistentServiceRangeVerification
{
    internal readonly record struct CellFormat(bool Bold, int FillColorIndex, string? NumberFormat);
    internal readonly record struct Geometry(double Left, double Top, double Width, double Height);

    internal static Geometry ReadGeometry(PersistentServiceWorkbookTestScope fixture,
        string sheetName, string address) =>
        fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                range = sheet.Range[address];
                return new Geometry(
                    Convert.ToDouble(((object)range.Left), CultureInfo.InvariantCulture),
                    Convert.ToDouble(((object)range.Top), CultureInfo.InvariantCulture),
                    Convert.ToDouble(((object)range.Width), CultureInfo.InvariantCulture),
                    Convert.ToDouble(((object)range.Height), CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });

    internal static CellFormat ReadFormat(PersistentServiceWorkbookTestScope fixture,
        string sheetName, string address) =>
        fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Font? font = null;
            Excel.Interior? interior = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range[address];
                font = range.Font;
                interior = range.Interior;
                return new CellFormat(
                    Convert.ToBoolean(((object)font.Bold), CultureInfo.InvariantCulture),
                    Convert.ToInt32(((object)interior.ColorIndex), CultureInfo.InvariantCulture),
                    Convert.ToString(((object)range.NumberFormat), CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
