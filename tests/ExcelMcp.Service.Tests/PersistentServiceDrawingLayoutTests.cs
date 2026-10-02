using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "DrawingLayout")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceDrawingLayoutTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void GroupUngroup_ReturnsMembersAndPreservesOtherObjects()
    {
        var sheet = CreateObjects();
        var grouped = _fixture.Send("drawing.group-objects", new
        {
            sheetName = sheet,
            objectNames = new List<string> { "First", "Second" },
            groupName = "Together"
        });
        using (var state = JsonDocument.Parse(grouped.Result!))
        {
            var group = Assert.Single(state.RootElement.GetProperty("drawingObjects").EnumerateArray());
            Assert.Equal("Together", group.GetProperty("name").GetString());
            Assert.Equal("Group", group.GetProperty("kind").GetString());
            Assert.Equal(2, group.GetProperty("children").GetArrayLength());
        }
        _fixture.Send("drawing.ungroup-object", new { sheetName = sheet, objectName = "Together" });
        var read = _fixture.Send("drawing.list-objects", new { sheetName = sheet });
        using var list = JsonDocument.Parse(read.Result!);
        Assert.Equal(3, list.RootElement.GetProperty("drawingObjects").GetArrayLength());
    }

    [Fact]
    public void AlignSelectedObjects_PreservesUnselectedGeometry()
    {
        var sheet = CreateObjects();
        _fixture.Send("drawing.align-objects", new
        {
            sheetName = sheet,
            objectNames = new List<string> { "First", "Second" },
            alignment = "Left"
        });
        var first = Read(sheet, "First");
        var second = Read(sheet, "Second");
        var third = Read(sheet, "Third");
        Assert.Equal(first.Left, second.Left, 2);
        Assert.Equal(200d, third.Left, 2);
        Assert.Equal(90d, third.Top, 2);
    }

    [Theory]
    [InlineData("Left")]
    [InlineData("Center")]
    [InlineData("Right")]
    [InlineData("Top")]
    [InlineData("Middle")]
    [InlineData("Bottom")]
    public void Alignment_AllNativeModesUseSelectedExtent(string alignment)
    {
        var sheet = CreateObjects();
        _fixture.Send("drawing.update-object", new { sheetName = sheet, objectName = "Second", width = 60d, height = 50d });
        var response = _fixture.Send("drawing.align-objects", new
        {
            sheetName = sheet,
            objectNames = new List<string> { "First", "Second" },
            alignment
        });
        using var document = JsonDocument.Parse(response.Result!);
        var objects = document.RootElement.GetProperty("drawingObjects").EnumerateArray().ToArray();
        Assert.Equal(2, objects.Length);
        static double Edge(JsonElement item, string mode) => mode switch
        {
            "Left" => item.GetProperty("left").GetDouble(),
            "Center" => item.GetProperty("left").GetDouble() + item.GetProperty("width").GetDouble() / 2,
            "Right" => item.GetProperty("left").GetDouble() + item.GetProperty("width").GetDouble(),
            "Top" => item.GetProperty("top").GetDouble(),
            "Middle" => item.GetProperty("top").GetDouble() + item.GetProperty("height").GetDouble() / 2,
            "Bottom" => item.GetProperty("top").GetDouble() + item.GetProperty("height").GetDouble(),
            _ => throw new ArgumentException("Unknown alignment.", nameof(mode))
        };
        Assert.Equal(Edge(objects[0], alignment), Edge(objects[1], alignment), 2);
    }

    [Theory]
    [InlineData("Horizontal", "left", "width")]
    [InlineData("Vertical", "top", "height")]
    public void Distribution_UsesEqualGapsWithDifferentObjectSizes(string distribution, string position, string size)
    {
        var sheet = CreateObjects();
        _fixture.Send("drawing.update-object", new { sheetName = sheet, objectName = "Second", width = 60d, height = 10d });
        var response = _fixture.Send("drawing.distribute-objects", new
        {
            sheetName = sheet,
            objectNames = new List<string> { "First", "Second", "Third" },
            distribution
        });
        using var document = JsonDocument.Parse(response.Result!);
        var objects = document.RootElement.GetProperty("drawingObjects").EnumerateArray()
            .OrderBy(item => item.GetProperty(position).GetDouble()).ToArray();
        Assert.Equal(3, objects.Length);
        var firstGap = objects[1].GetProperty(position).GetDouble() - objects[0].GetProperty(position).GetDouble() - objects[0].GetProperty(size).GetDouble();
        var secondGap = objects[2].GetProperty(position).GetDouble() - objects[1].GetProperty(position).GetDouble() - objects[1].GetProperty(size).GetDouble();
        Assert.Equal(firstGap, secondGap, 2);
        Assert.Equal(20d, objects[0].GetProperty(position).GetDouble(), 2);
        Assert.Equal(distribution == "Horizontal" ? 200d : 90d, objects[2].GetProperty(position).GetDouble(), 2);
    }

    [Theory]
    [InlineData("BringToFront", 3)]
    [InlineData("BringForward", 2)]
    [InlineData("SendToBack", 1)]
    [InlineData("SendBackward", 2)]
    public void ZOrder_ReturnsActualNativePosition(string zOrder, int expected)
    {
        var sheet = CreateObjects();
        var name = zOrder.StartsWith("Bring", StringComparison.Ordinal) ? "First" : "Third";
        var response = _fixture.Send("drawing.set-z-order", new { sheetName = sheet, objectName = name, zOrder });
        using var document = JsonDocument.Parse(response.Result!);
        var result = Assert.Single(document.RootElement.GetProperty("drawingObjects").EnumerateArray());
        Assert.Equal(expected, result.GetProperty("zOrderPosition").GetInt32());
    }

    [Fact]
    public void DuplicateGroup_ReturnsNewIdentityAndNativeMembership()
    {
        var sheet = CreateObjects();
        _fixture.Send("drawing.update-object", new { sheetName = sheet, objectName = "First", text = "Keep me", fillColor = "#70AD47" });
        _fixture.Send("drawing.group-objects", new { sheetName = sheet, objectNames = new List<string> { "First", "Second" }, groupName = "Inner" });
        _fixture.Send("drawing.group-objects", new { sheetName = sheet, objectNames = new List<string> { "Inner", "Third" }, groupName = "Outer" });
        var before = Read(sheet, "Outer");
        var original = _fixture.Send("drawing.get-object", new { sheetName = sheet, objectName = "Outer" });
        using var source = JsonDocument.Parse(original.Result!);
        var sourceMembers = source.RootElement.GetProperty("drawingObject").GetProperty("children");
        var response = _fixture.Send("drawing.duplicate-object", new { sheetName = sheet, objectName = "Outer", newName = "Copy", offsetLeft = 15d, offsetTop = -5d });
        using var document = JsonDocument.Parse(response.Result!);
        var copy = Assert.Single(document.RootElement.GetProperty("drawingObjects").EnumerateArray());
        Assert.Equal("Copy", copy.GetProperty("name").GetString());
        Assert.Equal(before.Left + 15, copy.GetProperty("left").GetDouble(), 2);
        Assert.Equal(before.Top - 5, copy.GetProperty("top").GetDouble(), 2);
        Assert.Equal(sourceMembers.GetArrayLength(), copy.GetProperty("children").GetArrayLength());
        Assert.Contains(copy.GetProperty("children").EnumerateArray(), item => item.TryGetProperty("text", out var text) && text.GetString() == "Keep me" && item.GetProperty("fillColor").GetString() == "#70AD47");
        Assert.Equal(before, Read(sheet, "Outer"));
        var read = _fixture.Send("drawing.get-object", new { sheetName = sheet, objectName = "Copy" });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.Equal(sourceMembers.GetArrayLength(), state.RootElement.GetProperty("drawingObject").GetProperty("children").GetArrayLength());
    }

    [Fact]
    public void Layout_SafeFormsControlAndTextBoxRoundTrip()
    {
        var sheet = CreateObjects();
        _fixture.Send("drawing.add-form-control", new { sheetName = sheet, controlType = "CheckBox", name = "Check", linkedCell = $"{sheet}!$J$1" });
        _fixture.Send("drawing.add-text-box", new { sheetName = sheet, name = "Note", text = "A note" });
        var response = _fixture.Send("drawing.group-objects", new { sheetName = sheet, objectNames = new List<string> { "Check", "Note" } });
        using var group = JsonDocument.Parse(response.Result!);
        var detail = Assert.Single(group.RootElement.GetProperty("drawingObjects").EnumerateArray());
        Assert.Equal(2, detail.GetProperty("children").GetArrayLength());
        _fixture.Send("drawing.ungroup-object", new { sheetName = sheet, objectName = detail.GetProperty("name").GetString() });
        var before = Read(sheet, "Check");
        var copied = _fixture.Send("drawing.duplicate-object", new { sheetName = sheet, objectName = "Check" });
        using var copy = JsonDocument.Parse(copied.Result!);
        var duplicate = Assert.Single(copy.RootElement.GetProperty("drawingObjects").EnumerateArray());
        Assert.Equal("FormControl", duplicate.GetProperty("kind").GetString());
        Assert.Equal(before.Left + 10, duplicate.GetProperty("left").GetDouble(), 2);
        Assert.Equal(before.Top + 10, duplicate.GetProperty("top").GetDouble(), 2);
    }

    [Theory]
    [InlineData("group-objects", """{"objectNames":["First","Missing"]}""")]
    [InlineData("group-objects", """{"objectNames":["First","first"]}""")]
    [InlineData("group-objects", """{"objectNames":["First","Second"],"groupName":"Third"}""")]
    [InlineData("group-objects", """{"objectNames":["First","Second"],"groupName":""}""")]
    [InlineData("group-objects", """{"objectNames":[]}""")]
    [InlineData("group-objects", """{"objectNames":["First"]}""")]
    [InlineData("align-objects", """{"objectNames":["First","Missing"],"alignment":"Left"}""")]
    [InlineData("align-objects", """{"objectNames":["First","Second"],"alignment":"Unknown"}""")]
    [InlineData("distribute-objects", """{"objectNames":["First","Second"],"distribution":"Horizontal"}""")]
    [InlineData("distribute-objects", """{"objectNames":["First","Second","Third"],"distribution":"Unknown"}""")]
    [InlineData("duplicate-object", """{"objectName":"First","newName":"Second"}""")]
    [InlineData("set-z-order", """{"objectName":"First","zOrder":"Unknown"}""")]
    [InlineData("ungroup-object", """{"objectName":"First"}""")]
    public async Task InvalidLayout_PreservesAllOriginalObjects(string action, string args)
    {
        var sheet = CreateObjects();
        var before = _fixture.Send("drawing.list-objects", new { sheetName = sheet }).Result;
        using var input = JsonDocument.Parse(args);
        var values = input.RootElement.EnumerateObject().ToDictionary(property => property.Name, property => (object?)property.Value);
        values["sheetName"] = sheet;
        var response = await _fixture.SendForFailureAsync($"drawing.{action}", values);
        Assert.False(response.Success);
        Assert.False(string.IsNullOrEmpty(response.ErrorMessage));
        Assert.Equal(before, _fixture.Send("drawing.list-objects", new { sheetName = sheet }).Result);
    }

    [Theory]
    [InlineData("group-objects", """{"objectNames":["First","Second"]}""")]
    [InlineData("align-objects", """{"objectNames":["First","Second"],"alignment":"Left"}""")]
    [InlineData("distribute-objects", """{"objectNames":["First","Second","Third"],"distribution":"Horizontal"}""")]
    [InlineData("duplicate-object", """{"objectName":"First"}""")]
    [InlineData("set-z-order", """{"objectName":"First","zOrder":"BringToFront"}""")]
    public async Task ProtectedDrawingObjects_BlockLayoutBeforeMutation(string action, string args)
    {
        var sheet = CreateObjects();
        var before = _fixture.Send("drawing.list-objects", new { sheetName = sheet }).Result;
        _fixture.Send("sheet.set-protection", new { sheetName = sheet, isProtected = true });
        using var input = JsonDocument.Parse(args);
        var values = input.RootElement.EnumerateObject().ToDictionary(property => property.Name, property => (object?)property.Value);
        values["sheetName"] = sheet;
        var response = await _fixture.SendForFailureAsync($"drawing.{action}", values);
        Assert.False(response.Success);
        Assert.Contains("protected", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        _fixture.Send("sheet.set-protection", new { sheetName = sheet, isProtected = false });
        Assert.Equal(before, _fixture.Send("drawing.list-objects", new { sheetName = sheet }).Result);
    }

    [Fact]
    public async Task Duplicate_RejectsMacroAssignmentsBeforeCreatingAnyCopy()
    {
        var sheet = CreateObjects();
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? worksheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                worksheet = (Excel.Worksheet)sheets[sheet];
                shapes = worksheet.Shapes;
                shape = shapes.Item("First");
                shape.OnAction = "NeverRunThis";
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref worksheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var response = await _fixture.SendForFailureAsync("drawing.duplicate-object", new { sheetName = sheet, objectName = "First" });
        Assert.False(response.Success);
        Assert.Contains("macro", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var list = _fixture.Send("drawing.list-objects", new { sheetName = sheet });
        using var document = JsonDocument.Parse(list.Result!);
        Assert.Equal(3, document.RootElement.GetProperty("drawingObjects").GetArrayLength());
    }

    private string CreateObjects()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        foreach (var (name, left, top) in new[] { ("First", 20d, 20d), ("Second", 100d, 70d), ("Third", 200d, 90d) })
            _fixture.Send("drawing.add-shape", new { sheetName = sheet, name, left, top, width = 40d, height = 30d });
        return sheet;
    }

    private (double Left, double Top) Read(string sheet, string name)
    {
        var response = _fixture.Send("drawing.get-object", new { sheetName = sheet, objectName = name });
        using var result = JsonDocument.Parse(response.Result!);
        var drawing = result.RootElement.GetProperty("drawingObject");
        return (drawing.GetProperty("left").GetDouble(), drawing.GetProperty("top").GetDouble());
    }
}
