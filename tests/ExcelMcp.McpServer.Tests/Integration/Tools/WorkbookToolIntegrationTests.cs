using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "false")]
public sealed class WorkbookToolIntegrationTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task DocumentProperty_SetAndGet_UsesGeneratedSnakeCaseContract()
    {
        const string sessionId = "recording-session";
        var setCall = await _fixture.CallToolAsync(
            "workbook",
            new Dictionary<string, object?>
            {
                ["action"] = "set-document-property",
                ["workbook_session_id"] = sessionId,
                ["property_name"] = "AutomationTag",
                ["value"] = "mcp-value",
                ["scope"] = "custom"
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "workbook.set-document-property",
            """{"propertyName":"AutomationTag","value":"mcp-value","scope":"custom"}""");

        using (var args = RecordingToolTest.ParseArgs(
            setCall.Request,
            "workbook.set-document-property",
            sessionId))
        {
            var root = args.RootElement;
            Assert.Equal("AutomationTag", root.GetProperty("propertyName").GetString());
            Assert.Equal("mcp-value", root.GetProperty("value").GetString());
            Assert.Equal("custom", root.GetProperty("scope").GetString());
        }

        var getCall = await _fixture.CallToolAsync(
            "workbook_read",
            new Dictionary<string, object?>
            {
                ["action"] = "get-document-property",
                ["workbook_session_id"] = sessionId,
                ["property_name"] = "AutomationTag",
                ["scope"] = "custom"
            },
            RecordingToolTest.Success(
                """{"success":true,"property":{"name":"AutomationTag","value":"mcp-value","scope":"custom"}}"""),
            "workbook.get-document-property",
            """{"propertyName":"AutomationTag","scope":"custom"}""");

        using var result = JsonDocument.Parse(getCall.JsonResult);
        var property = result.RootElement.GetProperty("property");
        Assert.Equal("AutomationTag", property.GetProperty("name").GetString());
        Assert.Equal("mcp-value", property.GetProperty("value").GetString());
        Assert.Equal("custom", property.GetProperty("scope").GetString());
    }

    [Theory]
    [InlineData("xlsx")]
    [InlineData("xlsm")]
    [InlineData("xlsb")]
    [InlineData("xls")]
    public async Task SaveAs_MapsEverySupportedFormat(string format)
    {
        const string sessionId = "recording-session";
        var targetPath = $@"C:\adapter-tests\saved.{format}";
        var saveCall = await _fixture.CallToolAsync(
            "workbook",
            new Dictionary<string, object?>
            {
                ["action"] = "save-as",
                ["workbook_session_id"] = sessionId,
                ["target_path"] = targetPath,
                ["format"] = format
            },
            RecordingToolTest.Success(
                $$$"""{"success":true,"fullName":"{{{targetPath.Replace("\\", "\\\\", StringComparison.Ordinal)}}}"}"""),
            "workbook.save-as",
            $$$"""{"targetPath":"{{{targetPath.Replace("\\", "\\\\", StringComparison.Ordinal)}}}","format":"{{{format}}}"}""");

        using (var args = RecordingToolTest.ParseArgs(
            saveCall.Request,
            "workbook.save-as",
            sessionId))
        {
            Assert.Equal(targetPath, args.RootElement.GetProperty("targetPath").GetString());
            Assert.Equal(format, args.RootElement.GetProperty("format").GetString());
        }
        using var result = JsonDocument.Parse(saveCall.JsonResult);
        Assert.Equal(targetPath, result.RootElement.GetProperty("fullName").GetString());
    }
}
