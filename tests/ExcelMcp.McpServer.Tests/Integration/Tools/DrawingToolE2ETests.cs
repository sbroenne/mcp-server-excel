using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Black-box MCP protocol coverage for worksheet drawing objects and sparklines.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Drawing")]
[Trait("RequiresExcel", "true")]
public sealed class DrawingToolE2ETests : McpIntegrationTestBase
{
    private readonly string _workbookPath;
    private string? _sessionId;

    public DrawingToolE2ETests(ITestOutputHelper output)
        : base(output, "DrawingToolE2EClient")
    {
        var tempDirectory = CreateTempDirectory("DrawingToolE2E");
        _workbookPath = Path.Join(tempDirectory, "DrawingTool.xlsx");
    }

    protected override async Task InitializeTestAsync()
    {
        _sessionId = await CreateWorkbookSessionAsync(_workbookPath);
    }

    [Fact]
    public async Task DrawingTool_ObjectAndSparklineLifecycle_SucceedsViaMcpProtocol()
    {
        var imagePath = CreateTestPng();
        var imageJson = await CallDrawingAsync("add-image", new()
        {
            ["sheet_name"] = "Sheet1",
            ["image_path"] = imagePath,
            ["name"] = "McpImage",
            ["left"] = 250,
            ["top"] = 100,
            ["width"] = 80,
            ["height"] = 60
        });
        AssertSuccess(imageJson, "drawing.add-image");

        var shapeJson = await CallDrawingAsync("add-shape", new()
        {
            ["sheet_name"] = "Sheet1",
            ["shape_type"] = "rounded-rectangle",
            ["name"] = "McpStatus",
            ["left"] = 20,
            ["top"] = 20,
            ["width"] = 180,
            ["height"] = 60,
            ["text"] = "Pending",
            ["fill_color"] = "#4472C4",
            ["line_color"] = "#203864",
            ["line_weight"] = 2
        });
        AssertSuccess(shapeJson, "drawing.add-shape");

        var textBoxJson = await CallDrawingAsync("add-text-box", new()
        {
            ["sheet_name"] = "Sheet1",
            ["text"] = "MCP note",
            ["name"] = "McpNote",
            ["left"] = 20,
            ["top"] = 100
        });
        AssertSuccess(textBoxJson, "drawing.add-text-box");

        var connectorJson = await CallDrawingAsync("add-connector", new()
        {
            ["sheet_name"] = "Sheet1",
            ["connector_type"] = "straight",
            ["begin_x"] = 40,
            ["begin_y"] = 180,
            ["end_x"] = 220,
            ["end_y"] = 180,
            ["name"] = "McpConnector"
        });
        AssertSuccess(connectorJson, "drawing.add-connector");

        var controlJson = await CallDrawingAsync("add-form-control", new()
        {
            ["sheet_name"] = "Sheet1",
            ["control_type"] = "check-box",
            ["name"] = "McpApproval",
            ["left"] = 250,
            ["top"] = 25,
            ["text"] = "Approved",
            ["linked_cell"] = "Sheet1!$J$2"
        });
        AssertSuccess(controlJson, "drawing.add-form-control");

        var updateJson = await CallDrawingAsync("update-object", new()
        {
            ["sheet_name"] = "Sheet1",
            ["object_name"] = "McpStatus",
            ["text"] = "Complete",
            ["fill_color"] = "#70AD47",
            ["rotation"] = 4
        });
        AssertSuccess(updateJson, "drawing.update-object");
        using (var updateDocument = JsonDocument.Parse(updateJson))
        {
            var drawingObject = updateDocument.RootElement.GetProperty("drawingObject");
            Assert.Equal("Complete", drawingObject.GetProperty("text").GetString());
            Assert.Equal("#70AD47", drawingObject.GetProperty("fillColor").GetString());
        }

        var getJson = await CallDrawingAsync("get-object", new()
        {
            ["sheet_name"] = "Sheet1",
            ["object_name"] = "McpStatus"
        });
        AssertSuccess(getJson, "drawing.get-object");

        var listJson = await CallDrawingAsync("list-objects", new()
        {
            ["sheet_name"] = "Sheet1"
        });
        AssertSuccess(listJson, "drawing.list-objects");
        using (var listDocument = JsonDocument.Parse(listJson))
        {
            Assert.Equal(5, listDocument.RootElement.GetProperty("drawingObjects").GetArrayLength());
        }

        var valuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["path"] = _workbookPath,
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "B2:E3",
            ["values"] = new List<List<object?>>
            {
                new() { 1, 3, 2, 5 },
                new() { 5, 2, 4, 1 }
            }
        });
        AssertSuccess(valuesJson, "range.set-values");

        var sparklineJson = await CallDrawingAsync("add-sparkline", new()
        {
            ["sheet_name"] = "Sheet1",
            ["source_range"] = "B2:E2",
            ["location_range"] = "F2",
            ["sparkline_type"] = "line",
            ["line_color"] = "#4472C4",
            ["show_markers"] = true
        });
        AssertSuccess(sparklineJson, "drawing.add-sparkline");

        var getSparklineJson = await CallDrawingAsync("get-sparkline", new()
        {
            ["sheet_name"] = "Sheet1",
            ["location_range"] = "F2"
        });
        AssertSuccess(getSparklineJson, "drawing.get-sparkline");

        var updateSparklineJson = await CallDrawingAsync("update-sparkline", new()
        {
            ["sheet_name"] = "Sheet1",
            ["location_range"] = "F2",
            ["source_range"] = "B3:E3",
            ["sparkline_type"] = "column",
            ["line_color"] = "#ED7D31",
            ["show_markers"] = false
        });
        AssertSuccess(updateSparklineJson, "drawing.update-sparkline");

        var listSparklinesJson = await CallDrawingAsync("list-sparklines", new()
        {
            ["sheet_name"] = "Sheet1"
        });
        AssertSuccess(listSparklinesJson, "drawing.list-sparklines");
        using (var sparklineDocument = JsonDocument.Parse(listSparklinesJson))
        {
            Assert.Single(sparklineDocument.RootElement.GetProperty("sparklines").EnumerateArray());
        }

        AssertSuccess(await CallDrawingAsync("delete-sparkline", new()
        {
            ["sheet_name"] = "Sheet1",
            ["location_range"] = "F2"
        }), "drawing.delete-sparkline");

        AssertSuccess(await CallDrawingAsync("delete-object", new()
        {
            ["sheet_name"] = "Sheet1",
            ["object_name"] = "McpStatus"
        }), "drawing.delete-object");
    }

    private Task<string> CallDrawingAsync(string action, Dictionary<string, object?> arguments)
    {
        arguments["action"] = action;
        arguments["path"] = _workbookPath;
        arguments["session_id"] = _sessionId;
        return CallToolAsync("drawing", arguments);
    }

    private string CreateTestPng()
    {
        const string onePixelPng =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
        var path = Path.Join(Path.GetDirectoryName(_workbookPath), $"Drawing_{Guid.NewGuid():N}.png");
        File.WriteAllBytes(path, Convert.FromBase64String(onePixelPng));
        return path;
    }
}
