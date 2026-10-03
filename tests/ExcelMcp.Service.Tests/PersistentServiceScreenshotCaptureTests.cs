using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Screenshot;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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
        AssertCapturedRange(result, sheetName);
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
        Assert.Contains("px", result.Message);
        AssertCapturedRange(result, sheetName);
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
        AssertCapturedRange(result, sheetName);
    }

    [Fact]
    public void CaptureSheet_EmbeddedChartBeyondUsedCells_IncludesChartPixels()
    {
        var batch = _fixture.BatchToken;
        var sheetName = PrepareSheet(batch, addChart: true);
        MoveAndMarkChart(sheetName);

        var result = _screenshotCommands.CaptureSheet(
            batch,
            sheetName,
            ScreenshotQuality.High);

        Assert.True(result.Success, result.ErrorMessage);
        AssertImageContainsChartMarker(result.ImageBase64);
        AssertCapturedRange(result, sheetName);
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
        AssertCapturedRange(result, sheetName);
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
        AssertCapturedRange(result, sheetName);
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
        AssertCapturedRange(result, sheetName);
    }

    private string PrepareSheet(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        bool addChart = false)
    {
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_commands.SetValues(
            batch,
            sheetName,
            "A1:B5",
            [
                ["Region", "Sales"],
                ["North", 45000],
                ["South", 38000],
                ["East", 51000],
                ["West", 42000]
            ]));
        Assert.True(_fixture.Send("rangeformat.format",
            new
            {
                sheetName,
                rangeAddresses = (string[])["A1:B5"],
                formatOptions = new { fillColor = "#28B4DC" }
            }).Success);

        if (addChart)
        {
            _fixture.ExecuteRawVerification((ctx, ct) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? source = null;
                Excel.ChartObjects? chartObjects = null;
                Excel.ChartObject? chartObject = null;
                Excel.Chart? chart = null;
                try
                {
                    sheets = ctx.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets[sheetName];
                    source = sheet.Range["A1:B5"];
                    chartObjects = (Excel.ChartObjects)sheet.ChartObjects();
                    chartObject = chartObjects.Add(150, 100, 400, 250);
                    chart = chartObject.Chart;
                    chart.SetSourceData(source);
                    chart.ChartType = Excel.XlChartType.xlColumnClustered;
                }
                finally
                {
                    ComUtilities.Release(ref chart);
                    ComUtilities.Release(ref chartObject);
                    ComUtilities.Release(ref chartObjects);
                    ComUtilities.Release(ref source);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });
        }

        return sheetName;
    }

    private void AssertCapturedRange(ScreenshotResult result, string sheetName)
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(sheetName, result.SheetName);
        Assert.NotNull(result.ImageBase64);
        using var stream = new MemoryStream(Convert.FromBase64String(result.ImageBase64));
        using var image = new Bitmap(stream);
        Assert.Equal(result.Width, image.Width);
        Assert.Equal(result.Height, image.Height);
        var matching = 0;
        for (var y = 0; y < image.Height; y += Math.Max(1, image.Height / 200))
        {
            for (var x = 0; x < image.Width; x += Math.Max(1, image.Width / 200))
            {
                var pixel = image.GetPixel(x, y);
                if (Math.Abs(pixel.R - 40) < 25 && Math.Abs(pixel.G - 180) < 25 && Math.Abs(pixel.B - 220) < 25)
                    matching++;
            }
        }
        Assert.True(matching >= 10, "The captured image is missing the target range's colored marker.");
        var values = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:B5");
        Assert.True(values.Success, values.ErrorMessage);
        Assert.Equal(["Region", "North", "South", "East", "West"],
            values.Values.Select(row => row[0]?.ToString()));
        Assert.Equal([45000, 38000, 51000, 42000],
            values.Values.Skip(1).Select(row => Convert.ToInt32(row[1], System.Globalization.CultureInfo.InvariantCulture)));
    }

    private void MoveAndMarkChart(string sheetName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ChartObjects? chartObjects = null;
            Excel.ChartObject? chartObject = null;
            Excel.Chart? chart = null;
            Excel.ChartArea? chartArea = null;
            Excel.Interior? chartInterior = null;
            Excel.Range? targetCell = null;

            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                chartObjects = (Excel.ChartObjects)sheet.ChartObjects();
                chartObject = (Excel.ChartObject)chartObjects.Item(1);
                targetCell = sheet.Range["BA1"];
                chartObject.Left = Convert.ToDouble(targetCell.Left, System.Globalization.CultureInfo.InvariantCulture);
                chartObject.Top = Convert.ToDouble(targetCell.Top, System.Globalization.CultureInfo.InvariantCulture);
                chart = chartObject.Chart;
                chartArea = chart.ChartArea;
                chartInterior = chartArea.Interior;
                chartInterior.Color = ColorTranslator.ToOle(Color.Magenta);
            }
            finally
            {
                ComUtilities.Release(ref targetCell);
                ComUtilities.Release(ref chartInterior);
                ComUtilities.Release(ref chartArea);
                ComUtilities.Release(ref chart);
                ComUtilities.Release(ref chartObject);
                ComUtilities.Release(ref chartObjects);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private static void AssertImageContainsChartMarker(string? base64)
    {
        Assert.NotNull(base64);
        using var stream = new MemoryStream(Convert.FromBase64String(base64));
        using var bitmap = new Bitmap(stream);

        int markerPixels = 0;
        int stepX = Math.Max(1, bitmap.Width / 200);
        int stepY = Math.Max(1, bitmap.Height / 200);

        for (int y = 0; y < bitmap.Height; y += stepY)
        {
            for (int x = 0; x < bitmap.Width; x += stepX)
            {
                Color pixel = bitmap.GetPixel(x, y);
                if (pixel.R > 220 && pixel.G < 80 && pixel.B > 220)
                {
                    markerPixels++;
                }
            }
        }

        Assert.True(markerPixels > 0, "Expected screenshot to contain the chart's magenta marker.");
    }

    private void Activate(string sheetName) =>
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                sheet.Activate();
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
