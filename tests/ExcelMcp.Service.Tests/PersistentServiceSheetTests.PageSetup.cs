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
        RequireSuccess(setResult);

        var getResult = _sheetCommands.GetPageSetup(batch, sheetName);
        RequireSuccess(getResult);
        Assert.Equal("landscape", getResult.Orientation);
        Assert.Equal(1, getResult.FitToPagesWide);
        Assert.Equal(2, getResult.FitToPagesTall);
        Assert.False(getResult.CenterHorizontally);
        Assert.True(getResult.CenterVertically);
        Assert.Equal(
            new NativePageSetupState("landscape", false, 1, 2, false, true),
            ReadNativePageSetup(sheetName));
    }

    [Fact]
    public void SetPageSetup_InvalidOrientation_PreservesExistingSettings()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var setup = _sheetCommands.SetPageSetup(batch, sheetName, "landscape", 1, 2, false, true);
        RequireSuccess(setup);
        var before = _sheetCommands.GetPageSetup(batch, sheetName);
        RequireSuccess(before);
        Assert.Equal("landscape", before.Orientation);
        Assert.Equal(1, before.FitToPagesWide);
        Assert.Equal(2, before.FitToPagesTall);
        Assert.False(before.CenterHorizontally);
        Assert.True(before.CenterVertically);
        var nativeBefore = ReadNativePageSetup(sheetName);
        Assert.Equal(new NativePageSetupState("landscape", false, 1, 2, false, true), nativeBefore);

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.SetPageSetup(batch, sheetName, "diagonal", 3, 4, true, false));

        Assert.Contains("Unsupported orientation 'diagonal'", exception.Message, StringComparison.Ordinal);
        var after = _sheetCommands.GetPageSetup(batch, sheetName);
        RequireSuccess(after);
        Assert.Equal(before.Orientation, after.Orientation);
        Assert.Equal(before.FitToPagesWide, after.FitToPagesWide);
        Assert.Equal(before.FitToPagesTall, after.FitToPagesTall);
        Assert.Equal(before.CenterHorizontally, after.CenterHorizontally);
        Assert.Equal(before.CenterVertically, after.CenterVertically);
        Assert.Equal(nativeBefore, ReadNativePageSetup(sheetName));
    }

    [Fact]
    public void GetPageSetup_AutomaticScaling_ReportsNativeZoomAndDefaults()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var result = _sheetCommands.GetPageSetup(_fixture.BatchToken, sheetName);

        RequireSuccess(result);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal("portrait", result.Orientation);
        Assert.Null(result.FitToPagesWide);
        Assert.Null(result.FitToPagesTall);
        Assert.False(result.CenterHorizontally);
        Assert.False(result.CenterVertically);
        Assert.Equal(
            new NativePageSetupState("portrait", true, null, null, false, false),
            ReadNativePageSetup(sheetName));
    }

    private NativePageSetupState ReadNativePageSetup(string sheetName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.PageSetup? pageSetup = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                pageSetup = sheet.PageSetup;
                var zoomEnabled = pageSetup.Zoom is bool zoom
                    ? zoom
                    : Convert.ToDouble(pageSetup.Zoom, System.Globalization.CultureInfo.InvariantCulture) != 0;
                int? fitToPagesWide = zoomEnabled
                    ? null
                    : Convert.ToInt32(pageSetup.FitToPagesWide, System.Globalization.CultureInfo.InvariantCulture);
                int? fitToPagesTall = zoomEnabled
                    ? null
                    : Convert.ToInt32(pageSetup.FitToPagesTall, System.Globalization.CultureInfo.InvariantCulture);
                return new NativePageSetupState(
                    pageSetup.Orientation == Excel.XlPageOrientation.xlLandscape ? "landscape" : "portrait",
                    zoomEnabled,
                    fitToPagesWide,
                    fitToPagesTall,
                    pageSetup.CenterHorizontally,
                    pageSetup.CenterVertically);
            }
            finally
            {
                ComUtilities.Release(ref pageSetup);
                ComUtilities.Release(ref sheet);
            }
        });

    private readonly record struct NativePageSetupState(
        string Orientation,
        bool ZoomEnabled,
        int? FitToPagesWide,
        int? FitToPagesTall,
        bool CenterHorizontally,
        bool CenterVertically);
}
