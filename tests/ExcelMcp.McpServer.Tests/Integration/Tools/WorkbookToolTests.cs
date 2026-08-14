using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "true")]
public sealed class WorkbookToolTests : McpIntegrationTestBase
{
    public WorkbookToolTests(ITestOutputHelper output)
        : base(output, "WorkbookClient")
    {
    }

    [Fact]
    public async Task Workbook_SetProtection_RoundsTripThroughMcp()
    {
        var tempDir = CreateTempDirectory("WorkbookProtection");
        var workbookPath = Path.Combine(tempDir, "workbook-protection.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);

        var protectJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "set-protection",
            ["session_id"] = sessionId,
            ["is_protected"] = true
        });
        AssertSuccess(protectJson, "workbook.set-protection");

        using var protectDoc = JsonDocument.Parse(protectJson);
        Assert.True(protectDoc.RootElement.GetProperty("success").GetBoolean());

        var getProtectionJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "get-protection",
            ["session_id"] = sessionId
        });
        AssertSuccess(getProtectionJson, "workbook.get-protection");

        using var getProtectionDoc = JsonDocument.Parse(getProtectionJson);
        Assert.True(getProtectionDoc.RootElement.GetProperty("isProtected").GetBoolean());

        var unprotectJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "set-protection",
            ["session_id"] = sessionId,
            ["is_protected"] = false
        });
        AssertSuccess(unprotectJson, "workbook.set-protection(false)");

        await TryCloseSessionAsync(sessionId, save: true);
    }

    [Fact]
    public async Task Workbook_SetViewOptions_RoundsTripThroughMcp()
    {
        var tempDir = CreateTempDirectory("WorkbookViewOptions");
        var workbookPath = Path.Combine(tempDir, "workbook-view-options.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);

        var setJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "set-view-options",
            ["session_id"] = sessionId,
            ["display_gridlines"] = false,
            ["display_headings"] = true
        });
        AssertSuccess(setJson, "workbook.set-view-options");

        var getJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "get-view-options",
            ["session_id"] = sessionId
        });
        AssertSuccess(getJson, "workbook.get-view-options");

        using var getDoc = JsonDocument.Parse(getJson);
        Assert.False(getDoc.RootElement.GetProperty("displayGridlines").GetBoolean());
        Assert.True(getDoc.RootElement.GetProperty("displayHeadings").GetBoolean());

        await TryCloseSessionAsync(sessionId, save: true);
    }
}
