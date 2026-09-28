using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class WorksheetImageToolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task WorksheetStyle_AddImageAndCountImages_RoundsTripThroughMcp()
    {
        const string sessionId = "recording-session";
        const string imagePath = @"C:\adapter-tests\sample.png";
        var addCall = await _fixture.CallToolAsync(
            "worksheet_style",
            new Dictionary<string, object?>
            {
                ["action"] = "add-image",
                ["session_id"] = sessionId,
                ["sheet_name"] = "ImageSheet",
                ["image_path"] = imagePath,
                ["cell_address"] = "A1"
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "sheet.add-image",
            """{"sheetName":"ImageSheet","imagePath":"C:\\adapter-tests\\sample.png","cellAddress":"A1"}""");

        using (var args = RecordingToolTest.ParseArgs(
            addCall.Request,
            "sheet.add-image",
            sessionId))
        {
            Assert.Equal("ImageSheet", args.RootElement.GetProperty("sheetName").GetString());
            Assert.Equal(imagePath, args.RootElement.GetProperty("imagePath").GetString());
            Assert.Equal("A1", args.RootElement.GetProperty("cellAddress").GetString());
        }

        var countCall = await _fixture.CallToolAsync(
            "worksheet_style",
            new Dictionary<string, object?>
            {
                ["action"] = "get-image-count",
                ["session_id"] = sessionId,
                ["sheet_name"] = "ImageSheet"
            },
            RecordingToolTest.Success("""{"success":true,"imageCount":1}"""),
            "sheet.get-image-count",
            """{"sheetName":"ImageSheet"}""");

        Assert.Equal("sheet.get-image-count", countCall.Request.Command);
        using var result = JsonDocument.Parse(countCall.JsonResult);
        Assert.Equal(1, result.RootElement.GetProperty("imageCount").GetInt32());
    }
}
