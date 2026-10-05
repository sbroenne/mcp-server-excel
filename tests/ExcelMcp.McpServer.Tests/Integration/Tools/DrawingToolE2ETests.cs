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
    private static readonly string[] ObjectNames =
        ["McpApproval", "McpConnector", "McpImage", "McpNote", "McpStatus"];
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
        AssertGeometry(ReadObject(imageJson, "McpImage", "Image"), 250, 100, 80, 60);

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
        var initialShape = ReadObject(shapeJson, "McpStatus", "AutoShape");
        AssertGeometry(initialShape, 20, 20, 180, 60);
        Assert.Equal("RoundedRectangle", initialShape.GetProperty("shapeType").GetString());
        Assert.Equal("Pending", initialShape.GetProperty("text").GetString());
        Assert.Equal("#4472C4", initialShape.GetProperty("fillColor").GetString());
        Assert.Equal("#203864", initialShape.GetProperty("lineColor").GetString());
        Assert.Equal(2d, initialShape.GetProperty("lineWeight").GetDouble());

        var textBoxJson = await CallDrawingAsync("add-text-box", new()
        {
            ["sheet_name"] = "Sheet1",
            ["text"] = "MCP note",
            ["name"] = "McpNote",
            ["left"] = 20,
            ["top"] = 100
        });
        AssertSuccess(textBoxJson, "drawing.add-text-box");
        Assert.Equal("MCP note", ReadObject(textBoxJson, "McpNote", "TextBox").GetProperty("text").GetString());

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
        var connector = ReadObject(connectorJson, "McpConnector", "Connector");
        AssertGeometry(connector, 40, 180, 180, 0);
        Assert.Equal("Straight", connector.GetProperty("connectorType").GetString());

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
        var control = ReadObject(controlJson, "McpApproval", "FormControl");
        Assert.Equal("CheckBox", control.GetProperty("formControlType").GetString());
        Assert.Equal("Approved", control.GetProperty("text").GetString());
        Assert.Equal("Sheet1!$J$2", control.GetProperty("linkedCell").GetString());

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
            Assert.Equal(4d, drawingObject.GetProperty("rotation").GetDouble(), 2);
        }

        var getJson = await CallDrawingAsync("get-object", new()
        {
            ["sheet_name"] = "Sheet1",
            ["object_name"] = "McpStatus"
        });
        AssertSuccess(getJson, "drawing.get-object");
        var readShape = ReadObject(getJson, "McpStatus", "AutoShape");
        Assert.Equal("Complete", readShape.GetProperty("text").GetString());
        Assert.Equal("#70AD47", readShape.GetProperty("fillColor").GetString());
        Assert.Equal(4d, readShape.GetProperty("rotation").GetDouble(), 2);
        AssertGeometry(readShape, 20, 20, 180, 60);

        var listJson = await CallDrawingAsync("list-objects", new()
        {
            ["sheet_name"] = "Sheet1"
        });
        AssertSuccess(listJson, "drawing.list-objects");
        using (var listDocument = JsonDocument.Parse(listJson))
        {
            Assert.Equal(5, listDocument.RootElement.GetProperty("drawingObjects").GetArrayLength());
            Assert.Equal(ObjectNames, listDocument.RootElement.GetProperty("drawingObjects").EnumerateArray()
                .Select(item => item.GetProperty("name").GetString()).Order(StringComparer.Ordinal));
        }

        var valuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["workbook_session_id"] = _sessionId,
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
        AssertSparkline(sparklineJson, "B2:E2", "Line", "#4472C4", true);

        var getSparklineJson = await CallDrawingAsync("get-sparkline", new()
        {
            ["sheet_name"] = "Sheet1",
            ["location_range"] = "F2"
        });
        AssertSuccess(getSparklineJson, "drawing.get-sparkline");
        AssertSparkline(getSparklineJson, "B2:E2", "Line", "#4472C4", true);

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
        AssertSparkline(updateSparklineJson, "B3:E3", "Column", "#ED7D31", false);

        var listSparklinesJson = await CallDrawingAsync("list-sparklines", new()
        {
            ["sheet_name"] = "Sheet1"
        });
        AssertSuccess(listSparklinesJson, "drawing.list-sparklines");
        using (var sparklineDocument = JsonDocument.Parse(listSparklinesJson))
        {
            AssertSparklineInfo(Assert.Single(sparklineDocument.RootElement.GetProperty("sparklines").EnumerateArray()),
                "B3:E3", "Column", "#ED7D31", false);
        }

        AssertSuccess(await CallDrawingAsync("delete-sparkline", new()
        {
            ["sheet_name"] = "Sheet1",
            ["location_range"] = "F2"
        }), "drawing.delete-sparkline");
        var deletedSparklines = await CallDrawingAsync("list-sparklines", new() { ["sheet_name"] = "Sheet1" });
        AssertSuccess(deletedSparklines, "drawing.list-sparklines after deletion");
        using (var document = JsonDocument.Parse(deletedSparklines))
            Assert.Empty(document.RootElement.GetProperty("sparklines").EnumerateArray());

        AssertSuccess(await CallDrawingAsync("delete-object", new()
        {
            ["sheet_name"] = "Sheet1",
            ["object_name"] = "McpStatus"
        }), "drawing.delete-object");
        var remainingObjects = await CallDrawingAsync("list-objects", new() { ["sheet_name"] = "Sheet1" });
        AssertSuccess(remainingObjects, "drawing.list-objects after deletion");
        using (var document = JsonDocument.Parse(remainingObjects))
            Assert.Equal(ObjectNames.Take(4), document.RootElement.GetProperty("drawingObjects").EnumerateArray()
                .Select(item => item.GetProperty("name").GetString()).Order(StringComparer.Ordinal));
    }

    private static JsonElement ReadObject(string json, string name, string kind)
    {
        using var document = JsonDocument.Parse(json);
        var item = document.RootElement.GetProperty("drawingObject");
        Assert.Equal(name, item.GetProperty("name").GetString());
        Assert.Equal("Sheet1", item.GetProperty("sheetName").GetString());
        Assert.Equal(kind, item.GetProperty("kind").GetString());
        return item.Clone();
    }

    private static void AssertGeometry(JsonElement item, double left, double top, double width, double height)
    {
        Assert.Equal(left, item.GetProperty("left").GetDouble(), 2);
        Assert.Equal(top, item.GetProperty("top").GetDouble(), 2);
        Assert.Equal(width, item.GetProperty("width").GetDouble(), 2);
        Assert.Equal(height, item.GetProperty("height").GetDouble(), 2);
    }

    private static void AssertSparkline(string json, string source, string type, string color, bool markers)
    {
        using var document = JsonDocument.Parse(json);
        AssertSparklineInfo(document.RootElement.GetProperty("sparkline"), source, type, color, markers);
    }

    private static void AssertSparklineInfo(JsonElement item, string source, string type, string color, bool markers)
    {
        Assert.Equal("Sheet1", item.GetProperty("sheetName").GetString());
        Assert.Equal("F2", item.GetProperty("locationRange").GetString());
        Assert.Equal(source, item.GetProperty("sourceRange").GetString());
        Assert.Equal(type, item.GetProperty("sparklineType").GetString());
        Assert.Equal(color, item.GetProperty("lineColor").GetString());
        Assert.Equal(markers, item.GetProperty("showMarkers").GetBoolean());
    }

    private Task<string> CallDrawingAsync(string action, Dictionary<string, object?> arguments)
    {
        arguments["action"] = action;
        arguments["workbook_session_id"] = _sessionId;
        var toolName = action is "get-object" or "list-objects" or "get-sparkline" or "list-sparklines"
            ? "drawing_read"
            : "drawing";
        return CallToolAsync(toolName, arguments);
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
