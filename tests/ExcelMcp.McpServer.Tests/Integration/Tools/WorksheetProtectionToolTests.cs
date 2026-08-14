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
public sealed class WorksheetProtectionToolTests : McpIntegrationTestBase
{
    public WorksheetProtectionToolTests(ITestOutputHelper output)
        : base(output, "WorksheetProtectionClient")
    {
    }

    [Fact]
    public async Task WorksheetStyle_SetProtection_RoundsTripThroughMcp()
    {
        var tempDir = CreateTempDirectory("WorksheetProtection");
        var workbookPath = Path.Combine(tempDir, "protection.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await CreateWorksheetAsync(sessionId, "ProtectedSheet");

        var protectJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "set-protection",
            ["session_id"] = sessionId,
            ["sheet_name"] = "ProtectedSheet",
            ["is_protected"] = true
        });
        AssertSuccess(protectJson, "worksheet_style.set-protection");

        using var protectDoc = JsonDocument.Parse(protectJson);
        Assert.True(protectDoc.RootElement.GetProperty("success").GetBoolean());

        var getProtectionJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "get-protection",
            ["session_id"] = sessionId,
            ["sheet_name"] = "ProtectedSheet"
        });
        AssertSuccess(getProtectionJson, "worksheet_style.get-protection");

        using var getProtectionDoc = JsonDocument.Parse(getProtectionJson);
        Assert.True(getProtectionDoc.RootElement.GetProperty("isProtected").GetBoolean());

        var unprotectJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "set-protection",
            ["session_id"] = sessionId,
            ["sheet_name"] = "ProtectedSheet",
            ["is_protected"] = false
        });
        AssertSuccess(unprotectJson, "worksheet_style.set-protection(false)");

        using var unprotectDoc = JsonDocument.Parse(unprotectJson);
        Assert.True(unprotectDoc.RootElement.GetProperty("success").GetBoolean());

        await TryCloseSessionAsync(sessionId, save: true);
    }
}
