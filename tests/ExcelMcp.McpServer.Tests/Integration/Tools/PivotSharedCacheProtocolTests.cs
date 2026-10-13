using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "PivotTables")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PivotSharedCacheProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Creation_DispatchesSharingPairAndOmittedDefaults(bool table, bool share)
    {
        var action = table ? "create-from-table" : "create-from-range";
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["pivot_table_name"] = "Second",
            ["destination_sheet"] = "Summary",
            ["destination_cell"] = "H1"
        };
        var expected = new Dictionary<string, object?>
        {
            ["pivotTableName"] = "Second",
            ["destinationSheet"] = "Summary",
            ["destinationCell"] = "H1"
        };
        if (table)
        {
            arguments["table_name"] = "Sales";
            expected["tableName"] = "Sales";
        }
        else
        {
            arguments["source_sheet"] = "Data";
            arguments["source_range"] = "A1:C20";
            expected["sourceSheet"] = "Data";
            expected["sourceRange"] = "A1:C20";
        }
        if (share)
        {
            arguments["share_cache_with"] = "First";
            arguments["share_cache_with_sheet"] = "Other Summary";
            expected["shareCacheWith"] = "First";
            expected["shareCacheWithSheet"] = "Other Summary";
        }
        var call = await fixture.CallToolAsync("pivottable", arguments,
            RecordingToolTest.Success("""{"success":true}"""), $"pivottable.{action}",
            JsonSerializer.Serialize(expected, ServiceProtocol.JsonOptions));
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task Discovery_SharingInputsAreOptionalAndOnlyOnWorksheetCreation()
    {
        var pivot = Assert.Single(await fixture.ListToolsAsync(), tool => tool.Name == "pivottable");
        var properties = pivot.JsonSchema.GetProperty("properties");
        foreach (var name in new[] { "share_cache_with", "share_cache_with_sheet" })
        {
            var schema = properties.GetProperty(name);
            Assert.Contains("string", schema.GetProperty("type").ToString(), StringComparison.Ordinal);
            Assert.Contains("cache", schema.GetProperty("description").GetString(), StringComparison.OrdinalIgnoreCase);
            foreach (var action in new[] { "create-from-table", "create-from-range" })
            {
                var parameter = Assert.Single(McpActionContract.GetParameters("pivottable", action), item => item.Name == name);
                Assert.False(parameter.Required);
            }
            Assert.DoesNotContain(McpActionContract.GetParameters("pivottable", "create-from-datamodel"),
                item => item.Name == name);
            Assert.DoesNotContain(McpActionContract.GetParameters("pivottable_read", "read"), item => item.Name == name);
        }
    }

    [Fact]
    public async Task DataModelCreation_RejectsSharingBeforeDispatch()
    {
        var result = await fixture.CallResultWithoutDispatchAsync("pivottable", new Dictionary<string, object?>
        {
            ["action"] = "create-from-datamodel",
            ["workbook_session_id"] = "session-1",
            ["pivot_table_name"] = "Second",
            ["table_name"] = "Sales",
            ["destination_sheet"] = "Summary",
            ["destination_cell"] = "H1",
            ["share_cache_with"] = "First",
            ["share_cache_with_sheet"] = "Summary"
        });
        Assert.True(result.IsError);
        var text = result.Content.OfType<ModelContextProtocol.Protocol.TextContentBlock>().First().Text;
        Assert.Contains("share_cache_with", text, StringComparison.Ordinal);
    }
}
