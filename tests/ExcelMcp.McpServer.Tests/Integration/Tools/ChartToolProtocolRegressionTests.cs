// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end regressions for chart tool behavior through the MCP protocol.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Charts")]
[Trait("RequiresExcel", "true")]
public sealed class ChartToolProtocolRegressionTests : McpIntegrationTestBase
{
    private readonly string _tempDir;

    public ChartToolProtocolRegressionTests(ITestOutputHelper output)
        : base(output, "ChartToolProtocolRegressionClient")
    {
        _tempDir = CreateTempDirectory("ChartToolProtocolRegressionTests");
    }

    [Fact]
    public async Task ChartList_EmptyWorkbook_ReturnsStructuredEmptyList_AndSessionRemainsUsable()
    {
        var workbookPath = Path.Join(_tempDir, $"NoCharts_{Guid.NewGuid():N}.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);

        var listResult = await CallToolAsync("chart", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));

        using (var listJson = JsonDocument.Parse(listResult))
        {
            var root = listJson.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean(), $"chart list failed: {listResult}");
            Assert.True(root.TryGetProperty("charts", out var charts), $"chart list should return charts: {listResult}");
            Assert.Equal(JsonValueKind.Array, charts.ValueKind);
            Assert.Empty(charts.EnumerateArray());
        }

        var worksheetListResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(worksheetListResult, "worksheet list after chart list");

        await CloseSessionAsync(sessionId, save: false);
    }
}
