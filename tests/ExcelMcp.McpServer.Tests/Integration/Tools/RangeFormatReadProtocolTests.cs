using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeFormatReadProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("stored")]
    [InlineData("displayed")]
    [InlineData("both")]
    [InlineData(null)]
    public async Task GetFormat_MapsViewAndPreservesEveryCell(string? view)
    {
        var response = JsonSerializer.Serialize(new
        {
            success = true,
            cellCount = 64,
            cells = Enumerable.Range(1, 64).Select(row => new { address = $"$A${row}" }).ToArray()
        });
        Dictionary<string, object?> arguments = new()
        {
            ["action"] = "get-format",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:A64"
        };
        if (view is not null)
        {
            arguments["view"] = view;
        }
        Dictionary<string, object?> serviceArguments = new()
        {
            ["sheetName"] = "Sheet1",
            ["rangeAddress"] = "A1:A64"
        };
        if (view is not null)
        {
            serviceArguments["view"] = view;
        }
        var expectedArgs = JsonSerializer.Serialize(serviceArguments, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_format", arguments,
            RecordingToolTest.Success(response), "rangeformat.get-format", expectedArgs);
        Assert.False(call.Result.IsError);
        using var output = JsonDocument.Parse(call.JsonResult);
        Assert.Equal(64, output.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(64, output.RootElement.GetProperty("cells").GetArrayLength());
    }

    [Fact]
    public async Task Discovery_AdvertisesCompleteStoredAndDisplayedReads()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range_format");
        Assert.Contains("get-format", tool.Description, StringComparison.Ordinal);
        Assert.Contains("displayed", tool.Description, StringComparison.Ordinal);
        Assert.True(tool.JsonSchema.GetProperty("properties").TryGetProperty("view", out _));
    }

    [Theory]
    [InlineData("unknown")]
    [InlineData("99")]
    public async Task GetFormat_InvalidViewDoesNotDispatch(string view)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("range_format", new()
        {
            ["action"] = "get-format",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["view"] = view
        });
        Assert.True(result.IsError);
    }
}
