using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "FineFormatting")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class FineFormattingProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task Format_ForwardsSelectedBordersAndExplicitFalse()
    {
        var options = new CellFormatOptions
        {
            Bold = false,
            FontThemeColor = 5,
            IndentLevel = 2,
            Borders = [new() { Position = CellBorderPosition.DiagonalUp, LineStyle = "dash", ThemeColor = 6 }]
        };
        var call = await fixture.CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format",
            ["session_id"] = "session-1",
            ["sheet_name"] = "Data",
            ["range_addresses"] = (string[])["A1:B2", "D1:E2"],
            ["format_options"] = new
            {
                bold = false,
                fontThemeColor = 5,
                indentLevel = 2,
                borders = new[] { new { position = "DiagonalUp", lineStyle = "dash", themeColor = 6 } }
            }
        }, RecordingToolTest.Success("""{"success":true}"""), "rangeformat.format",
            JsonSerializer.Serialize(new
            {
                sheetName = "Data",
                rangeAddresses = (string[])["A1:B2", "D1:E2"],
                formatOptions = options
            }, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_AdvertisesTypedFormattingAndNoObsoleteActions()
    {
        var tools = await fixture.ListToolsAsync();
        var tool = Assert.Single(tools, item => item.Name == "range_format");
        Assert.Contains("format_options", tool.Description, StringComparison.Ordinal);
        Assert.Contains("diagonal", tool.Description, StringComparison.Ordinal);
        var properties = tool.JsonSchema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("format_options", out _));
        Assert.DoesNotContain(properties.GetProperty("action").GetProperty("enum").EnumerateArray(),
            value => value.GetString() is "format-range" or "format-ranges");
    }
}
