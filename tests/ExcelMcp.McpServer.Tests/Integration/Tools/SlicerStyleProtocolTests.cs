using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Slicer")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SlicerStyleProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Theory]
    [InlineData("create-style")]
    [InlineData("apply-style")]
    [InlineData("list-styles")]
    public async Task Styles_MapNamesOptionsAndResults(string action)
    {
        var arguments = new Dictionary<string, object?>
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1"
        };
        var expected = new Dictionary<string, object?>();
        if (action != "list-styles")
        {
            arguments["style_name"] = "Brand";
            expected["styleName"] = "Brand";
        }
        if (action == "apply-style")
        {
            arguments["slicer_name"] = "Regions";
            expected["slicerName"] = "Regions";
        }
        if (action == "create-style")
        {
            const string options = """{"header":{"bold":false,"fillColor":"#123456","borderColor":"#654321","fontColor":"#FFFFFF"}}""";
            arguments["source_style_name"] = "SlicerStyleLight1";
            arguments["slicer_style_options"] = JsonSerializer.Deserialize<JsonElement>(options);
            expected["sourceStyleName"] = "SlicerStyleLight1";
            expected["slicerStyleOptions"] = JsonSerializer.Deserialize<JsonElement>(options);
        }
        const string response = """{"success":true,"style":{"name":"Brand"},"slicer":{"name":"Regions","style":"Brand"},"styles":[{"name":"Brand","builtIn":false}]}""";
        var call = await fixture.CallToolAsync(action == "list-styles" ? "slicer_read" : "slicer", arguments, RecordingToolTest.Success(response),
            $"slicer.{action}", action == "list-styles" ? null : JsonSerializer.Serialize(expected));
        Assert.False(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.Equal("Brand", result.RootElement.GetProperty("style").GetProperty("name").GetString());
        Assert.Equal("Brand", result.RootElement.GetProperty("slicer").GetProperty("style").GetString());
        Assert.False(result.RootElement.GetProperty("styles")[0].GetProperty("builtIn").GetBoolean());
    }

    [Fact]
    public async Task CreateStyle_PreservesPartialFailureDetails()
    {
        const string message = "Configuring the slicer style failed after Excel created style 'Brand'. Style 'Brand' remains in the workbook and may be partially formatted.";
        var call = await fixture.CallToolAsync("slicer", new()
        {
            ["action"] = "create-style",
            ["workbook_session_id"] = "session-1",
            ["style_name"] = "Brand",
            ["source_style_name"] = "SlicerStyleLight1",
            ["slicer_style_options"] = new { }
        }, new ServiceResponse
        {
            Success = false,
            ErrorCategory = "InvalidInput",
            ExceptionType = nameof(InvalidOperationException),
            ErrorMessage = message
        }, "slicer.create-style", """{"styleName":"Brand","sourceStyleName":"SlicerStyleLight1","slicerStyleOptions":{}}""");
        Assert.True(call.Result.IsError);
        using var result = JsonDocument.Parse(call.JsonResult);
        Assert.False(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(message, result.RootElement.GetProperty("errorMessage").GetString());
        Assert.Equal("InvalidInput", result.RootElement.GetProperty("errorCategory").GetString());
    }

    [Fact]
    public async Task Discovery_ExposesTypedOptionsAndReadOnlyListing()
    {
        var tools = await fixture.ListToolsAsync();
        var write = Assert.Single(tools, tool => tool.Name == "slicer");
        Assert.Contains("create-style", write.Description, StringComparison.Ordinal);
        Assert.Contains("apply-style", write.Description, StringComparison.Ordinal);
        var properties = write.JsonSchema.GetProperty("properties");
        Assert.Contains(properties.GetProperty("slicer_style_options").GetProperty("type").EnumerateArray(),
            type => type.GetString() == "object");
        var read = Assert.Single(tools, tool => tool.Name == "slicer_read");
        Assert.Contains("list-styles", read.Description, StringComparison.Ordinal);
        Assert.True(read.ProtocolTool.Annotations?.ReadOnlyHint == true);
    }
}
