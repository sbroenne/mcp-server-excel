using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void SetPageSetup_UpdatesSheetPageSetupProperties()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var setResult = _sheetCommands.SetPageSetup(
            batch,
            sheetName,
            "landscape",
            1,
            2,
            false,
            true);
        Assert.True(
            setResult.Success,
            $"Expected page setup to succeed but got error: {setResult.ErrorMessage}");

        var getResult = _sheetCommands.GetPageSetup(batch, sheetName);
        Assert.True(
            getResult.Success,
            $"Expected page setup read to succeed but got error: {getResult.ErrorMessage}");
        Assert.Equal("landscape", getResult.Orientation);
        Assert.Equal(1, getResult.FitToPagesWide);
        Assert.Equal(2, getResult.FitToPagesTall);
        Assert.False(getResult.CenterHorizontally);
        Assert.True(getResult.CenterVertically);
        Assert.False(IsPageSetupZoomEnabled(sheetName));
    }

    private bool IsPageSetupZoomEnabled(string sheetName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PageSetup? pageSetup = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                pageSetup = sheet.PageSetup;
                return pageSetup.Zoom is not bool value || value;
            }
            finally
            {
                ComUtilities.Release(ref pageSetup);
                ComUtilities.Release(ref sheet);
            }
        });
}
