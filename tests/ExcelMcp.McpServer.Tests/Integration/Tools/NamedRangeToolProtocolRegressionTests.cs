// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// End-to-end regressions for named range tool behavior through the MCP protocol.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public sealed class NamedRangeToolProtocolRegressionTests : McpIntegrationTestBase
{
    private readonly string _tempDir;

    public NamedRangeToolProtocolRegressionTests(ITestOutputHelper output)
        : base(output, "NamedRangeToolProtocolRegressionClient")
    {
        _tempDir = CreateTempDirectory("NamedRangeToolProtocolRegressionTests");
    }

    [Fact]
    public async Task NamedRangeList_EmptyWorkbook_ReturnsStructuredEmptyList_AndSessionRemainsUsable()
    {
        var workbookPath = Path.Join(_tempDir, $"NoNamedRanges_{Guid.NewGuid():N}.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);

        var listResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));

        using (var listJson = JsonDocument.Parse(listResult))
        {
            var root = listJson.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean(), $"namedrange list failed: {listResult}");
            Assert.True(root.TryGetProperty("namedRanges", out var namedRanges), $"namedrange list should return namedRanges: {listResult}");
            Assert.Equal(JsonValueKind.Array, namedRanges.ValueKind);
            Assert.Empty(namedRanges.EnumerateArray());
        }

        var worksheetListResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(worksheetListResult, "worksheet list after namedrange list");

        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task NamedRangeList_WorkbookWithNamedRange_ReturnsStructuredList_AndSessionRemainsUsable()
    {
        var workbookPath = Path.Join(_tempDir, $"OneNamedRange_{Guid.NewGuid():N}.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        var name = $"CsvFolder_{Guid.NewGuid():N}";

        var createResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = sessionId,
            ["name"] = name,
            ["reference"] = "Sheet1!$B$4"
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(createResult, "namedrange create setup");

        var writeResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "write",
            ["session_id"] = sessionId,
            ["name"] = name,
            ["value"] = "C:\\Data"
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(writeResult, "namedrange write setup");

        var listResult = await CallToolAsync("namedrange", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));

        using (var listJson = JsonDocument.Parse(listResult))
        {
            var root = listJson.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean(), $"namedrange list failed: {listResult}");
            var namedRanges = root.GetProperty("namedRanges").EnumerateArray().ToList();
            var matches = namedRanges.Where(range => range.GetProperty("name").GetString() == name).ToList();
            var listedRange = Assert.Single(matches);
            Assert.Contains("$B$4", listedRange.GetProperty("refersTo").GetString(), StringComparison.OrdinalIgnoreCase);
            Assert.Equal("C:\\Data", listedRange.GetProperty("value").GetString());
        }

        var worksheetListResult = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, TimeSpan.FromSeconds(30));
        AssertSuccess(worksheetListResult, "worksheet list after namedrange list");

        await CloseSessionAsync(sessionId, save: false);
    }
}
