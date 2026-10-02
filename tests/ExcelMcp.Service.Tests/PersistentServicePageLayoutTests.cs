using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PageLayout")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePageLayoutTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    private static readonly int[] ManualRows = [10, 20];
    private static readonly int[] ManualColumns = [4];

    [Fact]
    public void PageSetup_ReadsPrintScopeMarginsAndHeaders()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("sheet.set-page-setup", new
        {
            sheetName,
            orientation = "landscape",
            pageSetupOptions = new
            {
                printArea = "A1:D20",
                printTitleRows = "1:2",
                printTitleColumns = "A:B",
                leftMargin = 36d,
                rightMargin = 37d,
                topMargin = 38d,
                bottomMargin = 39d,
                headerMargin = 18d,
                footerMargin = 19d,
                leftHeader = "Left",
                centerHeader = "&P",
                rightHeader = "Right",
                leftFooter = "Foot",
                centerFooter = "&N",
                rightFooter = "End",
                paperSize = "xlPaperA4",
                pageOrder = "xlOverThenDown",
                printGridlines = true,
                printHeadings = true,
                zoomPercent = 85
            }
        });
        var response = _fixture.Send("sheet.get-page-setup", new { sheetName });
        using var result = JsonDocument.Parse(response.Result!);
        var state = result.RootElement;
        Assert.Equal("$A$1:$D$20", state.GetProperty("printArea").GetString());
        Assert.Equal("$1:$2", state.GetProperty("printTitleRows").GetString());
        Assert.Equal("$A:$B", state.GetProperty("printTitleColumns").GetString());
        Assert.Equal(36d, state.GetProperty("leftMargin").GetDouble(), 2);
        Assert.Equal(37d, state.GetProperty("rightMargin").GetDouble(), 2);
        Assert.Equal(38d, state.GetProperty("topMargin").GetDouble(), 2);
        Assert.Equal(39d, state.GetProperty("bottomMargin").GetDouble(), 2);
        Assert.Equal(18d, state.GetProperty("headerMargin").GetDouble(), 2);
        Assert.Equal(19d, state.GetProperty("footerMargin").GetDouble(), 2);
        Assert.Equal(85, state.GetProperty("zoomPercent").GetInt32());
        Assert.Equal("Left", state.GetProperty("leftHeader").GetString());
        Assert.Equal("&P", state.GetProperty("centerHeader").GetString());
        Assert.Equal("Right", state.GetProperty("rightHeader").GetString());
        Assert.Equal("Foot", state.GetProperty("leftFooter").GetString());
        Assert.Equal("&N", state.GetProperty("centerFooter").GetString());
        Assert.Equal("End", state.GetProperty("rightFooter").GetString());
        Assert.Equal("xlPaperA4", state.GetProperty("paperSize").GetString());
        Assert.Equal("xlOverThenDown", state.GetProperty("pageOrder").GetString());
        Assert.True(state.GetProperty("printGridlines").GetBoolean());
        Assert.True(state.GetProperty("printHeadings").GetBoolean());
    }

    [Fact]
    public void PageBreaks_ReplaceAndReadManualBreaks()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("range.set-values", new { sheetName, rangeAddress = "A1", values = new object[][] { ["Report"] } });
        _fixture.Send("sheet.set-page-setup", new
        {
            sheetName,
            orientation = "portrait",
            pageSetupOptions = new { printArea = "A1:H40" }
        });
        _fixture.Send("sheet.set-page-breaks", new
        {
            sheetName,
            pageBreakOptions = new { rows = ManualRows, columns = ManualColumns }
        });
        var response = _fixture.Send("sheet.get-page-breaks", new { sheetName });
        using var result = JsonDocument.Parse(response.Result!);
        var rows = result.RootElement.GetProperty("horizontal").EnumerateArray()
            .Where(item => item.GetProperty("isManual").GetBoolean()).Select(item => item.GetProperty("position").GetInt32()).ToArray();
        Assert.Equal([10, 20], rows);
        var columns = result.RootElement.GetProperty("vertical").EnumerateArray()
            .Where(item => item.GetProperty("isManual").GetBoolean()).Select(item => item.GetProperty("position").GetInt32()).ToArray();
        Assert.Equal([4], columns);
        _fixture.Send("sheet.set-page-breaks", new
        {
            sheetName,
            pageBreakOptions = new { rows = Array.Empty<int>(), columns = Array.Empty<int>() }
        });
        var cleared = _fixture.Send("sheet.get-page-breaks", new { sheetName });
        using var empty = JsonDocument.Parse(cleared.Result!);
        Assert.DoesNotContain(empty.RootElement.GetProperty("horizontal").EnumerateArray(),
            item => item.GetProperty("isManual").GetBoolean());
        Assert.DoesNotContain(empty.RootElement.GetProperty("vertical").EnumerateArray(),
            item => item.GetProperty("isManual").GetBoolean());
    }

    [Fact]
    public void PageSetup_ExplicitClearAndOmissionHaveDifferentEffects()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("sheet.set-page-setup", new
        {
            sheetName,
            orientation = "landscape",
            pageSetupOptions = new { printArea = "A1:B10,D1:E10", printTitleRows = "1:2", leftHeader = "Keep", rightFooter = "Remove", zoomPercent = 90 }
        });
        _fixture.Send("sheet.set-page-setup", new
        {
            sheetName,
            pageSetupOptions = new { printArea = "", printTitleRows = "", rightFooter = "", printGridlines = true }
        });
        var response = _fixture.Send("sheet.get-page-setup", new { sheetName });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.Equal("", result.RootElement.GetProperty("printArea").GetString());
        Assert.Equal("", result.RootElement.GetProperty("printTitleRows").GetString());
        Assert.Equal("", result.RootElement.GetProperty("rightFooter").GetString());
        Assert.Equal("Keep", result.RootElement.GetProperty("leftHeader").GetString());
        Assert.Equal("landscape", result.RootElement.GetProperty("orientation").GetString());
        Assert.Equal(90, result.RootElement.GetProperty("zoomPercent").GetInt32());
    }

    [Fact]
    public void PageSetup_FitAndFixedZoomSwitchNativeModes()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("sheet.set-page-setup", new
        {
            sheetName,
            fitToPagesWide = 1,
            fitToPagesTall = 0
        });
        var response = _fixture.Send("sheet.get-page-setup", new { sheetName });
        using (var result = JsonDocument.Parse(response.Result!))
        {
            Assert.Equal(JsonValueKind.Null, result.RootElement.GetProperty("zoomPercent").ValueKind);
            Assert.Equal(1, result.RootElement.GetProperty("fitToPagesWide").GetInt32());
            Assert.Equal(JsonValueKind.Null, result.RootElement.GetProperty("fitToPagesTall").ValueKind);
        }
        _fixture.Send("sheet.set-page-setup", new { sheetName, pageSetupOptions = new { zoomPercent = 100 } });
        var zoomRead = _fixture.Send("sheet.get-page-setup", new { sheetName });
        using var zoom = JsonDocument.Parse(zoomRead.Result!);
        Assert.Equal(100, zoom.RootElement.GetProperty("zoomPercent").GetInt32());
        Assert.Equal(JsonValueKind.Null, zoom.RootElement.GetProperty("fitToPagesWide").ValueKind);
    }

    [Theory]
    [InlineData("""{"zoomPercent":9}""")]
    [InlineData("""{"zoomPercent":401}""")]
    [InlineData("""{"leftMargin":-1}""")]
    [InlineData("""{"paperSize":"not-paper"}""")]
    [InlineData("""{"paperSize":"9"}""")]
    [InlineData("""{"pageOrder":"not-order"}""")]
    [InlineData("""{"printTitleRows":"A1:B2"}""")]
    [InlineData("""{"printTitleColumns":"A1:B2"}""")]
    [InlineData("""{"firstPageNumber":-1}""")]
    [InlineData("""{"unknownField":true}""")]
    public async Task PageSetup_InvalidOptionsDoNotMutateOrientation(string options)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("sheet.set-page-setup", new { sheetName, orientation = "portrait" });
        using var document = JsonDocument.Parse(options);
        var response = await _fixture.SendForFailureAsync("sheet.set-page-setup", new
        {
            sheetName,
            orientation = "landscape",
            pageSetupOptions = document.RootElement
        });
        Assert.False(response.Success);
        var read = _fixture.Send("sheet.get-page-setup", new { sheetName });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.Equal("portrait", state.RootElement.GetProperty("orientation").GetString());
    }

    [Fact]
    public async Task PageSetup_ConflictingScalingIsRejectedBeforeWriting()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = await _fixture.SendForFailureAsync("sheet.set-page-setup", new
        {
            sheetName,
            fitToPagesWide = 1,
            pageSetupOptions = new { zoomPercent = 100 }
        });
        Assert.False(response.Success);
        Assert.Contains("conflicting", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("""{"rows":[1],"columns":[]}""")]
    [InlineData("""{"rows":[1048577],"columns":[]}""")]
    [InlineData("""{"rows":[],"columns":[16385]}""")]
    [InlineData("""{"rows":[10,10],"columns":[]}""")]
    [InlineData("""{"rows":[10]}""")]
    [InlineData("""{"rows":[],"columns":null}""")]
    public async Task PageBreaks_InvalidReplacementRetainsExistingBreaks(string options)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("range.set-values", new { sheetName, rangeAddress = "A1", values = new object[][] { ["Report"] } });
        _fixture.Send("sheet.set-page-setup", new { sheetName, pageSetupOptions = new { printArea = "A1:H40" } });
        _fixture.Send("sheet.set-page-breaks", new { sheetName, pageBreakOptions = new { rows = ManualRows, columns = ManualColumns } });
        using var document = JsonDocument.Parse(options);
        var response = await _fixture.SendForFailureAsync("sheet.set-page-breaks", new
        {
            sheetName,
            pageBreakOptions = document.RootElement
        });
        Assert.False(response.Success);
        var read = _fixture.Send("sheet.get-page-breaks", new { sheetName });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.Equal(ManualRows, state.RootElement.GetProperty("horizontal").EnumerateArray()
            .Where(item => item.GetProperty("isManual").GetBoolean()).Select(item => item.GetProperty("position").GetInt32()));
    }

    [Fact]
    public void PageSetup_AdditionalNativePrintingSettingsRoundTrip()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("sheet.set-page-setup", new
        {
            sheetName,
            pageSetupOptions = new
            {
                blackAndWhite = true,
                draft = true,
                firstPageNumber = 7,
                printComments = "xlPrintSheetEnd",
                printErrors = "xlPrintErrorsDash",
                scaleWithDocHeaderFooter = false,
                alignMarginsHeaderFooter = false
            }
        });
        var read = _fixture.Send("sheet.get-page-setup", new { sheetName });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.True(state.RootElement.GetProperty("blackAndWhite").GetBoolean());
        Assert.True(state.RootElement.GetProperty("draft").GetBoolean());
        Assert.Equal(7, state.RootElement.GetProperty("firstPageNumber").GetInt32());
        Assert.Equal("xlPrintSheetEnd", state.RootElement.GetProperty("printComments").GetString());
        Assert.Equal("xlPrintErrorsDash", state.RootElement.GetProperty("printErrors").GetString());
        Assert.False(state.RootElement.GetProperty("scaleWithDocHeaderFooter").GetBoolean());
        Assert.False(state.RootElement.GetProperty("alignMarginsHeaderFooter").GetBoolean());
        _fixture.Send("sheet.set-page-setup", new { sheetName, pageSetupOptions = new { firstPageNumber = 0 } });
        var automatic = _fixture.Send("sheet.get-page-setup", new { sheetName });
        using var auto = JsonDocument.Parse(automatic.Result!);
        Assert.Equal(0, auto.RootElement.GetProperty("firstPageNumber").GetInt32());
    }
}
