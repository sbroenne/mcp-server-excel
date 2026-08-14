using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "true")]
public sealed class WorksheetPageSetupToolTests : McpIntegrationTestBase
{
    public WorksheetPageSetupToolTests(ITestOutputHelper output)
        : base(output, "WorksheetPageSetupClient")
    {
    }

    [Fact]
    public async Task WorksheetStyle_SetPageSetup_RoundsTripThroughMcp()
    {
        var tempDir = CreateTempDirectory("WorksheetPageSetup");
        var workbookPath = Path.Combine(tempDir, "pagesetup.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await CreateWorksheetAsync(sessionId, "PageSetupSheet");

        var setJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "set-page-setup",
            ["session_id"] = sessionId,
            ["sheet_name"] = "PageSetupSheet",
            ["orientation"] = "landscape",
            ["fit_to_pages_wide"] = 1,
            ["fit_to_pages_tall"] = 2,
            ["center_horizontally"] = false,
            ["center_vertically"] = true
        });
        AssertSuccess(setJson, "worksheet_style.set-page-setup");

        var getJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "get-page-setup",
            ["session_id"] = sessionId,
            ["sheet_name"] = "PageSetupSheet"
        });
        AssertSuccess(getJson, "worksheet_style.get-page-setup");

        using var doc = JsonDocument.Parse(getJson);
        Assert.Equal("landscape", doc.RootElement.GetProperty("orientation").GetString());
        Assert.Equal(1, doc.RootElement.GetProperty("fitToPagesWide").GetInt32());
        Assert.Equal(2, doc.RootElement.GetProperty("fitToPagesTall").GetInt32());
        Assert.False(doc.RootElement.GetProperty("centerHorizontally").GetBoolean());
        Assert.True(doc.RootElement.GetProperty("centerVertically").GetBoolean());

        await TryCloseSessionAsync(sessionId, save: true);
    }

    [Fact]
    public async Task WorksheetStyle_GetPageSetup_AutomaticScalingReturnsNullFitValues()
    {
        var tempDir = CreateTempDirectory("WorksheetPageSetupAutomatic");
        var workbookPath = Path.Combine(tempDir, "pagesetup-automatic.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await CreateWorksheetAsync(sessionId, "AutomaticScale");

        var getJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "get-page-setup",
            ["session_id"] = sessionId,
            ["sheet_name"] = "AutomaticScale"
        });
        AssertSuccess(getJson, "worksheet_style.get-page-setup");

        using var doc = JsonDocument.Parse(getJson);
        Assert.Equal(JsonValueKind.Null, doc.RootElement.GetProperty("fitToPagesWide").ValueKind);
        Assert.Equal(JsonValueKind.Null, doc.RootElement.GetProperty("fitToPagesTall").ValueKind);

        await TryCloseSessionAsync(sessionId, save: false);
    }
}
