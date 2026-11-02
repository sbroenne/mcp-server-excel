using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Slicer")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceSlicerStyleTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Theory]
    [InlineData("wholeTable", "xlWholeTable")]
    [InlineData("header", "xlHeaderRow")]
    [InlineData("selectedItemWithData", "xlSlicerSelectedItemWithData")]
    [InlineData("selectedItemWithNoData", "xlSlicerSelectedItemWithNoData")]
    [InlineData("unselectedItemWithData", "xlSlicerUnselectedItemWithData")]
    [InlineData("unselectedItemWithNoData", "xlSlicerUnselectedItemWithNoData")]
    [InlineData("hoveredSelectedItemWithData", "xlSlicerHoveredSelectedItemWithData")]
    [InlineData("hoveredSelectedItemWithNoData", "xlSlicerHoveredSelectedItemWithNoData")]
    [InlineData("hoveredUnselectedItemWithData", "xlSlicerHoveredUnselectedItemWithData")]
    [InlineData("hoveredUnselectedItemWithNoData", "xlSlicerHoveredUnselectedItemWithNoData")]
    public void CreateStyle_MapsEveryElementAndPreservesSource(string friendlyName, string nativeName)
    {
        const string source = "SlicerStyleLight1";
        string name = $"Brand_{Guid.NewGuid():N}";
        var before = _fixture.Send("workbook.get-table-style", new { styleName = source });
        Assert.True(before.Success, before.ErrorMessage);
        var response = _fixture.Send("slicer.create-style", new
        {
            styleName = name,
            sourceStyleName = source,
            slicerStyleOptions = new Dictionary<string, object>
            {
                [friendlyName] = new { fillColor = "#123456", borderColor = "#654321", fontColor = "#234567", bold = false }
            }
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        var style = result.RootElement.GetProperty("style");
        Assert.Equal(name, style.GetProperty("name").GetString());
        Assert.False(style.GetProperty("builtIn").GetBoolean());
        Assert.True(style.GetProperty("showAsAvailableSlicerStyle").GetBoolean());
        var element = Element(result, nativeName);
        Assert.Equal("#123456", element.GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
        Assert.Equal("#234567", element.GetProperty("font").GetProperty("color").GetProperty("rgb").GetString());
        Assert.False(element.GetProperty("font").GetProperty("bold").GetBoolean());
        foreach (string edge in new[] { "xlEdgeLeft", "xlEdgeTop", "xlEdgeBottom", "xlEdgeRight" })
        {
            var border = Assert.Single(element.GetProperty("borders").EnumerateArray(),
                item => item.GetProperty("edge").GetString() == edge);
            Assert.Equal("#654321", border.GetProperty("color").GetProperty("rgb").GetString());
            Assert.Equal(1, border.GetProperty("lineStyle").GetInt32());
            Assert.Equal(2, border.GetProperty("weight").GetInt32());
        }
        using var original = JsonDocument.Parse(before.Result!);
        foreach (var untouched in original.RootElement.GetProperty("style").GetProperty("elements").EnumerateArray())
            if (untouched.GetProperty("nativeType").GetInt32() != element.GetProperty("nativeType").GetInt32())
                Assert.Equal(untouched.GetRawText(), Element(result, untouched.GetProperty("elementType").GetString()!).GetRawText());
        Assert.Equal(before.Result, _fixture.Send("workbook.get-table-style", new { styleName = source }).Result);
        var read = _fixture.Send("workbook.get-table-style", new { styleName = name });
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal(response.Result, read.Result);
    }

    [Fact]
    public async Task Styles_CopyCustomApplyAndPersistWithoutChangingLayout()
    {
        string name = $"Brand_{Guid.NewGuid():N}";
        CreateStyle(name, "SlicerStyleLight1", new { header = new { bold = true, fillColor = "#ABCDEF" } });
        string copy = name + "_Copy";
        CreateStyle(copy, name, new { });
        using var original = JsonDocument.Parse(_fixture.Send("workbook.get-table-style", new { styleName = name }).Result!);
        using var copied = JsonDocument.Parse(_fixture.Send("workbook.get-table-style", new { styleName = copy }).Result!);
        Assert.Equal(original.RootElement.GetProperty("style").GetProperty("elements").GetRawText(),
            copied.RootElement.GetProperty("style").GetProperty("elements").GetRawText());
        CreateTableSlicer();
        _fixture.Send("slicer.update-slicer", new { slicerName = "StyleRegions", slicerOptions = new { columnCount = 2, width = 240d, caption = "Keep caption" } });
        _fixture.Send("slicer.set-table-slicer-selection", new { slicerName = "StyleRegions", selectedItems = new List<string> { "North" } });
        using var before = State("StyleRegions");
        var applied = _fixture.Send("slicer.apply-style", new { slicerName = "StyleRegions", styleName = copy });
        Assert.True(applied.Success, applied.ErrorMessage);
        using var after = State("StyleRegions");
        AssertStyleOnly(before, after, copy);
        await _fixture.SaveAndReopenAsync();
        using var reopened = State("StyleRegions");
        AssertStyleOnly(before, reopened, copy);
        using var definition = JsonDocument.Parse(_fixture.Send("workbook.get-table-style", new { styleName = copy }).Result!);
        Assert.Equal("#ABCDEF", Element(definition, "xlHeaderRow").GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
        Assert.True(Element(definition, "xlHeaderRow").GetProperty("font").GetProperty("bold").GetBoolean());
        _fixture.Send("slicer.apply-style", new { slicerName = "StyleRegions", styleName = "SlicerStyleLight2" });
        using var builtin = State("StyleRegions");
        AssertStyleOnly(before, builtin, "SlicerStyleLight2");
    }

    [Fact]
    public void ListStyles_FiltersTheNativeCatalogue()
    {
        string name = $"Brand_{Guid.NewGuid():N}";
        CreateStyle(name, "SlicerStyleLight1", new { });
        using var all = JsonDocument.Parse(_fixture.Send("workbook.list-table-styles", new { }).Result!);
        var response = _fixture.Send("slicer.list-styles", new { });
        Assert.True(response.Success, response.ErrorMessage);
        using var list = JsonDocument.Parse(response.Result!);
        var actual = list.RootElement.GetProperty("styles").EnumerateArray().ToArray();
        Assert.Equal(all.RootElement.GetProperty("styles").EnumerateArray()
                .Where(style => style.GetProperty("showAsAvailableSlicerStyle").GetBoolean()).Select(style => style.GetRawText()),
            actual.Select(style => style.GetRawText()));
        Assert.Contains(actual, style => style.GetProperty("name").GetString() == name && !style.GetProperty("builtIn").GetBoolean());
        Assert.Contains(actual, style => style.GetProperty("name").GetString() == "SlicerStyleLight1" && style.GetProperty("builtIn").GetBoolean());
        Assert.DoesNotContain(actual, style => style.GetProperty("name").GetString() is "TableStyleMedium2" or "TimeSlicerStyleLight1");
    }

    [Theory]
    [InlineData("SlicerStyleLight1", """{"header":{"fillColor":"invalid"}}""")]
    [InlineData("SlicerStyleLight1", """{"header":{"borderColor":"invalid"}}""")]
    [InlineData("SlicerStyleLight1", """{"header":{"fontColor":"invalid"}}""")]
    [InlineData("SlicerStyleLight1", """{"unknown":{}}""")]
    [InlineData("SlicerStyleLight1", """{"header":{"unknown":true}}""")]
    [InlineData("MissingSourceStyle", "{}")]
    [InlineData("TableStyleMedium2", "{}")]
    public async Task CreateStyle_InvalidInputsLeaveCatalogueUnchanged(string source, string options)
    {
        var before = _fixture.Send("workbook.list-table-styles", new { });
        using var parsed = JsonDocument.Parse(options);
        var failure = await _fixture.SendForFailureAsync("slicer.create-style", new
        {
            styleName = $"Invalid_{Guid.NewGuid():N}",
            sourceStyleName = source,
            slicerStyleOptions = parsed.RootElement
        });
        Assert.False(failure.Success);
        Assert.False(string.IsNullOrWhiteSpace(failure.ErrorMessage));
        Assert.Equal(before.Result, _fixture.Send("workbook.list-table-styles", new { }).Result);
    }

    [Fact]
    public async Task CreateStyle_DuplicateNamePreservesDefinition()
    {
        string name = $"Brand_{Guid.NewGuid():N}";
        CreateStyle(name, "SlicerStyleLight1", new { header = new { bold = true } });
        var before = _fixture.Send("workbook.get-table-style", new { styleName = name });
        var failure = await _fixture.SendForFailureAsync("slicer.create-style", new { styleName = name, sourceStyleName = "SlicerStyleLight1", slicerStyleOptions = new { header = new { bold = false } } });
        Assert.False(failure.Success);
        Assert.Contains("already exists", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before.Result, _fixture.Send("workbook.get-table-style", new { styleName = name }).Result);
    }

    [Theory]
    [InlineData("MissingStyle")]
    [InlineData("TableStyleMedium2")]
    public async Task ApplyStyle_InvalidStylePreservesControl(string style)
    {
        CreateTableSlicer();
        using var before = State("StyleRegions");
        var failure = await _fixture.SendForFailureAsync("slicer.apply-style",
            new { slicerName = "StyleRegions", styleName = style });
        Assert.False(failure.Success);
        Assert.False(string.IsNullOrWhiteSpace(failure.ErrorMessage));
        using var after = State("StyleRegions");
        Assert.Equal(before.RootElement.GetProperty("slicer").GetRawText(), after.RootElement.GetProperty("slicer").GetRawText());
    }

    [Fact]
    public async Task ApplyStyle_MissingControlDoesNotChangeOtherControl()
    {
        CreateTableSlicer();
        using var before = State("StyleRegions");
        var failure = await _fixture.SendForFailureAsync("slicer.apply-style",
            new { slicerName = "MissingSlicer", styleName = "SlicerStyleLight2" });
        Assert.False(failure.Success);
        Assert.Contains("MissingSlicer", failure.ErrorMessage, StringComparison.Ordinal);
        using var after = State("StyleRegions");
        Assert.Equal(before.RootElement.GetProperty("slicer").GetRawText(), after.RootElement.GetProperty("slicer").GetRawText());
    }

    [Fact]
    public async Task ApplyStyle_ProtectedControlPreservesState()
    {
        CreateTableSlicer();
        using var before = State("StyleRegions");
        SetDrawingProtection(true);
        try
        {
            var failure = await _fixture.SendForFailureAsync("slicer.apply-style",
                new { slicerName = "StyleRegions", styleName = "SlicerStyleLight2" });
            Assert.False(failure.Success);
            Assert.Contains("Unprotect", failure.ErrorMessage, StringComparison.Ordinal);
            using var after = State("StyleRegions");
            Assert.Equal(before.RootElement.GetProperty("slicer").GetRawText(), after.RootElement.GetProperty("slicer").GetRawText());
        }
        finally { SetDrawingProtection(false); }
    }

    [Fact]
    public async Task ApplyStyle_TimelineIsRejectedWithoutChanges()
    {
        const string sheet = "StyleDates";
        _fixture.CreateNamedTestSheet(_fixture.BatchToken, sheet);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B4",
            [["Date", "Amount"], [new DateTime(2026, 1, 1), 10], [new DateTime(2026, 2, 1), 20], [new DateTime(2026, 3, 1), 30]]));
        _fixture.Send("pivottable.create-from-range", new { sourceSheet = sheet, sourceRange = "A1:B4", destinationSheet = sheet, destinationCell = "D1", pivotTableName = "StyleDatePivot" });
        _fixture.Send("pivottablefield.add-value-field", new { pivotTableName = "StyleDatePivot", fieldName = "Amount" });
        _fixture.Send("slicer.create-timeline", new { pivotTableName = "StyleDatePivot", fieldName = "Date", slicerName = "StyleTimeline", destinationSheet = sheet, position = "H2" });
        using var before = State("StyleTimeline");
        var failure = await _fixture.SendForFailureAsync("slicer.apply-style", new { slicerName = "StyleTimeline", styleName = "SlicerStyleLight1" });
        Assert.False(failure.Success);
        Assert.Contains("ordinary slicer", failure.ErrorMessage, StringComparison.Ordinal);
        using var after = State("StyleTimeline");
        Assert.Equal(before.RootElement.GetProperty("slicer").GetRawText(), after.RootElement.GetProperty("slicer").GetRawText());
    }

    private void CreateTableSlicer()
    {
        const string sheet = "StyleData";
        _fixture.CreateNamedTestSheet(_fixture.BatchToken, sheet);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B4",
            [["Region", "Amount"], ["North", 10], ["South", 20], ["North", 30]]));
        _fixture.Send("table.create", new { sheetName = sheet, tableName = "StyleSales", rangeAddress = "A1:B4" });
        _fixture.RegisterTableForCleanup("StyleSales");
        _fixture.Send("slicer.create-table-slicer", new { tableName = "StyleSales", columnName = "Region", slicerName = "StyleRegions", destinationSheet = sheet, position = "D2" });
    }

    private void SetDrawingProtection(bool protect) =>
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, "StyleData");
                Assert.NotNull(sheet);
                if (protect) sheet.Protect(DrawingObjects: true, Contents: false, Scenarios: false);
                else sheet.Unprotect();
            }
            finally { ComUtilities.Release(ref sheet); }
        });

    private void CreateStyle(string name, string source, object options)
    {
        var response = _fixture.Send("slicer.create-style", new { styleName = name, sourceStyleName = source, slicerStyleOptions = options });
        Assert.True(response.Success, response.ErrorMessage);
    }

    private JsonDocument State(string name)
    {
        var response = _fixture.Send("slicer.get-slicer", new { slicerName = name });
        Assert.True(response.Success, response.ErrorMessage);
        return JsonDocument.Parse(response.Result!);
    }

    internal static void AssertStyleOnly(JsonDocument before, JsonDocument after, string style)
    {
        var actual = after.RootElement.GetProperty("slicer");
        Assert.Equal(style, actual.GetProperty("style").GetString());
        foreach (var property in before.RootElement.GetProperty("slicer").EnumerateObject())
            if (property.Name != "style")
                Assert.Equal(property.Value.GetRawText(), actual.GetProperty(property.Name).GetRawText());
    }

    private static JsonElement Element(JsonDocument document, string name) =>
        Assert.Single(document.RootElement.GetProperty("style").GetProperty("elements").EnumerateArray(),
            element => element.GetProperty("elementType").GetString() == name);
}
