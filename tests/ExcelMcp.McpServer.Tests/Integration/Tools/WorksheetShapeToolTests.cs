using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class WorksheetShapeToolTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task WorksheetStyle_AddShapeAndCountShapes_RoundsTripThroughMcp()
    {
        const string sessionId = "recording-session";
        var addCall = await _fixture.CallToolAsync(
            "worksheet_style",
            new Dictionary<string, object?>
            {
                ["action"] = "add-shape",
                ["session_id"] = sessionId,
                ["sheet_name"] = "ShapeSheet",
                ["cell_address"] = "A1"
            },
            RecordingToolTest.Success("""{"success":true}"""),
            "sheet.add-shape",
            """{"sheetName":"ShapeSheet","cellAddress":"A1"}""");

        using (var args = RecordingToolTest.ParseArgs(
            addCall.Request,
            "sheet.add-shape",
            sessionId))
        {
            Assert.Equal("ShapeSheet", args.RootElement.GetProperty("sheetName").GetString());
            Assert.Equal("A1", args.RootElement.GetProperty("cellAddress").GetString());
        }

        var countCall = await _fixture.CallToolAsync(
            "worksheet_style",
            new Dictionary<string, object?>
            {
                ["action"] = "get-shape-count",
                ["session_id"] = sessionId,
                ["sheet_name"] = "ShapeSheet"
            },
            RecordingToolTest.Success("""{"success":true,"shapeCount":1}"""),
            "sheet.get-shape-count",
            """{"sheetName":"ShapeSheet"}""");

        Assert.Equal("sheet.get-shape-count", countCall.Request.Command);
        using var result = JsonDocument.Parse(countCall.JsonResult);
        Assert.Equal(1, result.RootElement.GetProperty("shapeCount").GetInt32());
    }
}
