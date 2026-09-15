using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Screenshot")]
public sealed partial class PersistentServiceScreenshotCaptureTests :
    PersistentServiceWorkbookTestBase,
    IClassFixture<PersistentServiceScreenshotFixture>
{
    private readonly IScreenshotCommands _screenshotCommands;
    private readonly ISheetStyleCommands _sheetCommands;

    public PersistentServiceScreenshotCaptureTests(
        PersistentServiceScreenshotFixture fixture) :
        base(fixture)
    {
        _screenshotCommands = fixture.CreateCommands<IScreenshotCommands>();
        _sheetCommands = fixture.CreateCommands<ISheetStyleCommands>();
    }

    [Fact]
    public void CaptureRange_SmallRange_ReturnsValidPng()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch);

        var result = _screenshotCommands.CaptureRange(
            batch,
            sheetName,
            "A1:B5",
            ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.ImageBase64);
        Assert.NotEmpty(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        var bytes = Convert.FromBase64String(result.ImageBase64);
        Assert.True(bytes.Length > 100);
        Assert.Equal([137, 80, 78, 71], bytes[..4]);
    }

    [Fact]
    public void CaptureRange_MediumQuality_ReturnsJpeg()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch);

        var result = _screenshotCommands.CaptureRange(
            batch,
            sheetName,
            "A1:B5");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.ImageBase64);
        Assert.NotEmpty(result.ImageBase64);
        Assert.Equal("image/jpeg", result.MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        var bytes = Convert.FromBase64String(result.ImageBase64);
        Assert.True(bytes.Length > 100);
        Assert.Equal(0xFF, bytes[0]);
        Assert.Equal(0xD8, bytes[1]);
    }

    [Fact]
    public void CaptureRange_AreaWithChart_ReturnsLargerImage()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch, addChart: true);

        var result = _screenshotCommands.CaptureRange(
            batch,
            sheetName,
            "A1:M20",
            ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        Assert.True(Convert.FromBase64String(result.ImageBase64).Length > 500);
    }

    [Fact]
    public void CaptureSheet_EmbeddedChart_ExpandsCaptureBeyondUsedCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch, addChart: true);

        var result = _screenshotCommands.CaptureSheet(
            batch,
            sheetName,
            ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotEqual("$A$1:$B$5", result.RangeAddress);
        Assert.True(result.Width >= 500, $"Expected chart-inclusive capture width but got {result.Width}px.");
        Assert.True(result.Height >= 300, $"Expected chart-inclusive capture height but got {result.Height}px.");
    }

    [Fact]
    public void CaptureSheet_NamedSheet_ReturnsValidPng()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch);

        var result = _screenshotCommands.CaptureSheet(
            batch,
            sheetName,
            ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.ImageBase64);
        Assert.NotEmpty(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        Assert.Equal(sheetName, result.SheetName);
    }

    [Fact]
    public void CaptureSheet_ActiveSheet_ReturnsValidPng()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch);
        Activate(sheetName);

        var result = _screenshotCommands.CaptureSheet(
            batch,
            quality: ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
    }

    [Fact]
    public void CaptureRange_DefaultRange_ReturnsValidPng()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch);
        Activate(sheetName);

        var result = _screenshotCommands.CaptureRange(
            batch,
            quality: ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.NotNull(result.ImageBase64);
        Assert.Equal("image/png", result.MimeType);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
    }

    [Fact]
    public void CaptureRange_MessageIncludesDimensions()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch);

        var result = _screenshotCommands.CaptureRange(
            batch,
            sheetName,
            "A1:B5");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains("px", result.Message);
    }

    private string PrepareSheet(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        bool addChart = false)
    {
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(
            batch,
            sheetName,
            "A1:B5",
            [
                ["Region", "Sales"],
                ["North", 45000],
                ["South", 38000],
                ["East", 51000],
                ["West", 42000]
            ]);

        if (addChart)
        {
            _fixture.ExecuteRawVerification((ctx, ct) =>
            {
                dynamic? sheet = null;
                dynamic? chartObjects = null;
                dynamic? chartObject = null;
                dynamic? chart = null;
                try
                {
                    sheet = ctx.Book.Worksheets[sheetName];
                    chartObjects = sheet.ChartObjects();
                    chartObject = chartObjects.Add(150, 100, 400, 250);
                    chart = chartObject.Chart;
                    chart.SetSourceData(sheet.Range["A1:B5"]);
                    chart.ChartType = 51;
                }
                finally
                {
                    ComUtilities.Release(ref chart);
                    ComUtilities.Release(ref chartObject);
                    ComUtilities.Release(ref chartObjects);
                    ComUtilities.Release(ref sheet);
                }
            });
        }

        return sheetName;
    }

    private void Activate(string sheetName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic? sheet = null;
            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                sheet.Activate();
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
}
