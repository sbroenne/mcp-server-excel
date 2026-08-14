using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "true")]
public sealed class WorksheetShapeToolTests : McpIntegrationTestBase
{
    public WorksheetShapeToolTests(ITestOutputHelper output)
        : base(output, "WorksheetShapeClient")
    {
    }

    [Fact]
    public async Task WorksheetStyle_AddShapeAndCountShapes_RoundsTripThroughMcp()
    {
        var tempDir = CreateTempDirectory("WorksheetShapes");
        var workbookPath = Path.Combine(tempDir, "shapes.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await CreateWorksheetAsync(sessionId, "ShapeSheet");

        var addShapeJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "add-shape",
            ["session_id"] = sessionId,
            ["sheet_name"] = "ShapeSheet",
            ["cell_address"] = "A1"
        });
        AssertSuccess(addShapeJson, "worksheet_style.add-shape");

        var getShapeCountJson = await CallToolAsync("worksheet_style", new Dictionary<string, object?>
        {
            ["action"] = "get-shape-count",
            ["session_id"] = sessionId,
            ["sheet_name"] = "ShapeSheet"
        });
        AssertSuccess(getShapeCountJson, "worksheet_style.get-shape-count");

        using var getShapeCountDoc = JsonDocument.Parse(getShapeCountJson);
        Assert.True(getShapeCountDoc.RootElement.GetProperty("shapeCount").GetInt32() > 0);

        await TryCloseSessionAsync(sessionId, save: true);
    }
}
