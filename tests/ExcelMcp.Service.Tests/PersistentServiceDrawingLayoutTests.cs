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
        _fixture.Send("drawing.update-object", new { sheetName = sheet, objectName = "First", text = "Grouped content", fillColor = "#70AD47" });
        _fixture.Send("drawing.update-object", new { sheetName = sheet, objectName = "Third", text = "Unselected content", fillColor = "#FF0000" });
        var before = ReadNativeGeometry(sheet);
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
            Assert.Equal(["First", "Second"],
                group.GetProperty("children").EnumerateArray().Select(child => child.GetProperty("name").GetString()).Order());
            var first = Assert.Single(group.GetProperty("children").EnumerateArray(),
                child => child.GetProperty("name").GetString() == "First");
            Assert.Equal("Grouped content", first.GetProperty("text").GetString());
            Assert.Equal("#70AD47", first.GetProperty("fillColor").GetString());
        }
        AssertGeometry(before["Third"], ReadNativeGeometry(sheet)["Third"]);
        AssertMaterial("Third", "Unselected content", "#FF0000");
        _fixture.Send("drawing.ungroup-object", new { sheetName = sheet, objectName = "Together" });
        var read = _fixture.Send("drawing.list-objects", new { sheetName = sheet });
        using var list = JsonDocument.Parse(read.Result!);
        Assert.Equal(3, list.RootElement.GetProperty("drawingObjects").GetArrayLength());
        Assert.Equal(before.Keys.Order(),
            list.RootElement.GetProperty("drawingObjects").EnumerateArray().Select(item => item.GetProperty("name").GetString()).Order());
        var after = ReadNativeGeometry(sheet);
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        foreach (var name in before.Keys)
            AssertGeometry(before[name], after[name]);
        AssertMaterial("First", "Grouped content", "#70AD47");
        AssertMaterial("Third", "Unselected content", "#FF0000");

        void AssertMaterial(string name, string text, string color)
        {
            using var read = JsonDocument.Parse(_fixture.Send("drawing.get-object", new { sheetName = sheet, objectName = name }).Result!);
            var item = read.RootElement.GetProperty("drawingObject");
            Assert.Equal(text, item.GetProperty("text").GetString());
            Assert.Equal(color, item.GetProperty("fillColor").GetString());
        }
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
    [InlineData("Left", 20d)]
    [InlineData("Center", 90d)]
    [InlineData("Right", 160d)]
    [InlineData("Top", 20d)]
    [InlineData("Middle", 70d)]
    [InlineData("Bottom", 120d)]
    public void Alignment_AllNativeModesUseSelectedExtent(string alignment, double expectedEdge)
    {
        var sheet = CreateObjects();
        _fixture.Send("drawing.update-object", new { sheetName = sheet, objectName = "Second", width = 60d, height = 50d });
        var before = ReadNativeGeometry(sheet);
        var response = _fixture.Send("drawing.align-objects", new
        {
            sheetName = sheet,
            objectNames = new List<string> { "First", "Second" },
            alignment
        });
        using var document = JsonDocument.Parse(response.Result!);
        var objects = document.RootElement.GetProperty("drawingObjects").EnumerateArray().ToArray();
        Assert.Equal(2, objects.Length);
        Assert.Equal(["First", "Second"], objects.Select(item => item.GetProperty("name").GetString()).Order());
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
        Assert.All(objects, item => Assert.Equal(expectedEdge, Edge(item, alignment), 2));
        var after = ReadNativeGeometry(sheet);
        bool horizontal = alignment is "Left" or "Center" or "Right";
        foreach (var name in new[] { "First", "Second" })
        {
            var original = before[name];
            var actual = after[name];
            Assert.Equal(original.Width, actual.Width, 2);
            Assert.Equal(original.Height, actual.Height, 2);
            Assert.Equal(horizontal ? original.Top : original.Left, horizontal ? actual.Top : actual.Left, 2);
            double nativeEdge = alignment switch
            {
                "Left" => actual.Left,
                "Center" => actual.Left + actual.Width / 2,
                "Right" => actual.Left + actual.Width,
                "Top" => actual.Top,
                "Middle" => actual.Top + actual.Height / 2,
                "Bottom" => actual.Top + actual.Height,
                _ => throw new ArgumentException("Unknown alignment.", nameof(alignment))
            };
            Assert.Equal(expectedEdge, nativeEdge, 2);
        }
        AssertGeometry(before["Third"], after["Third"]);
    }

    [Theory]
    [InlineData("Horizontal", "left", "width")]
    [InlineData("Vertical", "top", "height")]
    public void Distribution_UsesEqualGapsWithDifferentObjectSizes(string distribution, string position, string size)
    {
        var sheet = CreateObjects();
        _fixture.Send("drawing.update-object", new
        {
            sheetName = sheet,
            objectName = "Second",
            left = 130d,
            top = 75d,
            width = 60d,
            height = 10d
        });
        var before = ReadNativeGeometry(sheet);
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
        var after = ReadNativeGeometry(sheet);
        Assert.Equal(before.Keys.Order(), after.Keys.Order());
        bool horizontal = distribution == "Horizontal";
        double expectedGap = horizontal
            ? (before["Third"].Left - before["First"].Left - before["First"].Width - before["Second"].Width) / 2
            : (before["Third"].Top - before["First"].Top - before["First"].Height - before["Second"].Height) / 2;
        Assert.NotEqual(horizontal ? before["Second"].Left : before["Second"].Top,
            horizontal ? after["Second"].Left : after["Second"].Top);
        Assert.Equal(expectedGap, horizontal
            ? after["Second"].Left - after["First"].Left - after["First"].Width
            : after["Second"].Top - after["First"].Top - after["First"].Height, 2);
        Assert.Equal(expectedGap, horizontal
            ? after["Third"].Left - after["Second"].Left - after["Second"].Width
            : after["Third"].Top - after["Second"].Top - after["Second"].Height, 2);
        foreach (var name in before.Keys)
        {
            Assert.Equal(before[name].Width, after[name].Width, 2);
            Assert.Equal(before[name].Height, after[name].Height, 2);
            Assert.Equal(horizontal ? before[name].Top : before[name].Left,
                horizontal ? after[name].Top : after[name].Left, 2);
        }
        Assert.Equal(horizontal ? before["First"].Left : before["First"].Top,
            horizontal ? after["First"].Left : after["First"].Top, 2);
        Assert.Equal(horizontal ? before["Third"].Left : before["Third"].Top,
            horizontal ? after["Third"].Left : after["Third"].Top, 2);
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
        Assert.Equal(expected, ReadNativeOrderAndAction(sheet, name).Position);
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
        var failure = await Record.ExceptionAsync(async () =>
        {
            using var input = JsonDocument.Parse(args);
            var values = input.RootElement.EnumerateObject().ToDictionary(property => property.Name, property => (object?)property.Value);
            values["sheetName"] = sheet;
            var response = await _fixture.SendForFailureAsync($"drawing.{action}", values);
            Assert.False(response.Success);
            Assert.Contains("protected", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(before, _fixture.Send("drawing.list-objects", new { sheetName = sheet }).Result);
        });
        var cleanup = Record.Exception(() =>
            _fixture.Send("sheet.set-protection", new { sheetName = sheet, isProtected = false }));
        if (cleanup is not null)
            failure = PersistentServiceCleanupFailures.Combine(failure, cleanup);
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
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
        var before = _fixture.Send("drawing.list-objects", new { sheetName = sheet }).Result;
        var actionBefore = ReadNativeOrderAndAction(sheet, "First");
        var response = await _fixture.SendForFailureAsync("drawing.duplicate-object", new { sheetName = sheet, objectName = "First" });
        Assert.False(response.Success);
        Assert.Contains("macro", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var list = _fixture.Send("drawing.list-objects", new { sheetName = sheet });
        using var document = JsonDocument.Parse(list.Result!);
        Assert.Equal(3, document.RootElement.GetProperty("drawingObjects").GetArrayLength());
        Assert.Equal(before, list.Result);
        Assert.Equal(actionBefore, ReadNativeOrderAndAction(sheet, "First"));
    }

    [Theory]
    [InlineData("drawing.group-objects")]
    [InlineData("drawing.duplicate-object")]
    public async Task NameExcelRejects_IsRejectedBeforeChangingObjects(string command)
    {
        // Excel accepts drawing names up to 254 characters; 255 fails only after the group or copy exists.
        var sheet = CreateObjects();
        var name = new string('G', 255);
        var before = _fixture.Send("drawing.list-objects", new { sheetName = sheet }).Result;

        var response = command == "drawing.group-objects"
            ? await _fixture.SendForFailureAsync(command, new
            {
                sheetName = sheet,
                objectNames = new List<string> { "First", "Second" },
                groupName = name
            })
            : await _fixture.SendForFailureAsync(command, new { sheetName = sheet, objectName = "First", newName = name });

        Assert.False(response.Success);
        Assert.Contains("at most 254 characters", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(before, _fixture.Send("drawing.list-objects", new { sheetName = sheet }).Result);
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

    private (int Position, string Action) ReadNativeOrderAndAction(string sheetName, string name) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                shapes = sheet.Shapes;
                shape = shapes.Item(name);
                return (shape.ZOrderPosition, shape.OnAction);
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref sheet);
            }
        });

    private Dictionary<string, (double Left, double Top, double Width, double Height)> ReadNativeGeometry(string sheet) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? worksheet = null;
            Excel.Shapes? shapes = null;
            Excel.Shape? shape = null;
            try
            {
                worksheet = ComUtilities.FindSheet(context.Book, sheet);
                shapes = worksheet!.Shapes;
                var result = new Dictionary<string, (double Left, double Top, double Width, double Height)>(StringComparer.Ordinal);
                for (int index = 1; index <= shapes.Count; index++)
                {
                    shape = shapes.Item(index);
                    result.Add(shape.Name, (shape.Left, shape.Top, shape.Width, shape.Height));
                    ComUtilities.Release(ref shape);
                }
                return result;
            }
            finally
            {
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref shapes);
                ComUtilities.Release(ref worksheet);
            }
        });

    private static void AssertGeometry(
        (double Left, double Top, double Width, double Height) expected,
        (double Left, double Top, double Width, double Height) actual)
    {
        Assert.Equal(expected.Left, actual.Left, 2);
        Assert.Equal(expected.Top, actual.Top, 2);
        Assert.Equal(expected.Width, actual.Width, 2);
        Assert.Equal(expected.Height, actual.Height, 2);
    }
}
