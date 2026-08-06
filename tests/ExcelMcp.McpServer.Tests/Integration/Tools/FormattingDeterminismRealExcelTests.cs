using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "DeterministicFormatting")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class FormattingDeterminismRealExcelTests : McpIntegrationTestBase
{
    public FormattingDeterminismRealExcelTests(ITestOutputHelper output)
        : base(output, "FormattingDeterminismClient")
    {
    }

    protected override IReadOnlyList<string> ServerArguments =>
        ["--tool-profile", "copilot-compact"];

    [Fact]
    public async Task ReportOutlineAndFreeze_RoundTripThroughThePublicCompactMcpSurface()
    {
        var directory = CreateTempDirectory("FormattingDeterminism");
        var workbookPath = Path.Join(directory, "deterministic-formatting.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);

        var seed = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:D6",
            ["values"] = new List<List<object?>>
            {
                new() { "Quarterly Sales", null, null, null },
                new() { "Product", "Price", "Units", "Total" },
                new() { "Alpha", 10, 2, 20 },
                new() { "Beta", 12, 3, 36 },
                new() { "Gamma", 8, 4, 32 },
                new() { "Total", null, null, 88 },
            },
        });
        AssertSuccess(seed, "range.set-values seed");

        var applied = await CallToolAsync("layout", new Dictionary<string, object?>
        {
            ["action"] = "apply-report",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["title_range"] = "A1:D1",
            ["header_range"] = "A2:D2",
            ["body_range"] = "A3:D5",
            ["total_range"] = "A6:D6",
            ["preset"] = "professional",
            ["accent_color"] = "#1F4E78",
            ["auto_fit_columns"] = true,
        });
        using var appliedDocument = JsonDocument.Parse(applied);
        Assert.True(appliedDocument.RootElement.GetProperty("success").GetBoolean(), applied);
        Assert.Equal(64, appliedDocument.RootElement.GetProperty("fingerprint").GetString()!.Length);
        Assert.Equal(4, appliedDocument.RootElement.GetProperty("sections").GetArrayLength());

        var reportState = await CallToolAsync("layout", new Dictionary<string, object?>
        {
            ["action"] = "get-report",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["title_range"] = "A1:D1",
            ["header_range"] = "A2:D2",
            ["body_range"] = "A3:D5",
            ["total_range"] = "A6:D6",
        });
        using var reportStateDocument = JsonDocument.Parse(reportState);
        Assert.Equal(
            appliedDocument.RootElement.GetProperty("fingerprint").GetString(),
            reportStateDocument.RootElement.GetProperty("fingerprint").GetString());

        var outline = await CallToolAsync("layout", new Dictionary<string, object?>
        {
            ["action"] = "set-outline",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "3:5",
            ["axis"] = "row",
            ["level"] = 1,
            ["collapsed"] = true,
        });
        using var outlineDocument = JsonDocument.Parse(outline);
        Assert.True(outlineDocument.RootElement.GetProperty("success").GetBoolean(), outline);
        Assert.Equal(1, outlineDocument.RootElement.GetProperty("level").GetInt32());
        Assert.True(outlineDocument.RootElement.GetProperty("collapsed").GetBoolean());

        var readValues = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A3:D5",
        });
        using var valuesDocument = JsonDocument.Parse(readValues);
        Assert.Equal("Alpha", valuesDocument.RootElement.GetProperty("values")[0][0].GetString());
        Assert.Equal(36d, valuesDocument.RootElement.GetProperty("values")[1][3].GetDouble());

        await CloseSessionAsync(sessionId, save: false);
    }
}
