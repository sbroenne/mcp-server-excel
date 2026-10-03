using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeGetStyleTests
{
    [Fact]
    public void GetStyle_MixedStyles_PreservesOptionalFallbackWithoutChangingCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetStyle(batch, sheetName, "A1", "Good").Success);
        Assert.True(_commands.SetStyle(batch, sheetName, "A2", "Bad").Success);

        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            object? style = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range["A1:A2"];
                style = range.Style;
                Assert.True(style is null or DBNull,
                    "Excel must report no single style for the deliberately mixed range.");
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

        var result = _commands.GetStyle(batch, sheetName, "A1:A2");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal("Normal", result.StyleName);
        Assert.Equal("$A$1:$A$2", result.RangeAddress);
        Assert.Equal(sheetName, result.SheetName);

        var first = _commands.GetStyle(batch, sheetName, "A1");
        var second = _commands.GetStyle(batch, sheetName, "A2");
        Assert.True(first.Success, first.ErrorMessage);
        Assert.True(second.Success, second.ErrorMessage);
        Assert.Equal("Good", first.StyleName);
        Assert.Equal("Bad", second.StyleName);
    }
}
