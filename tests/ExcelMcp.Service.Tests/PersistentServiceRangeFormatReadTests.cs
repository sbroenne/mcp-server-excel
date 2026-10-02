using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceRangeFormatReadTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void GetFormat_ReturnsEveryCellAndPreservesDifferentFormatting()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1"],
            formatOptions = new
            {
                bold = true,
                fontColor = "#123456",
                fillColor = "#ABCDEF",
                horizontalAlignment = "right",
                wrapText = true
            }
        });
        _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["B1"],
            formatOptions = new
            {
                bold = false,
                fontSize = 18,
                fillColor = "#FFFFFF"
            }
        });
        Assert.True(_commands.SetNumberFormat(
            _fixture.BatchToken, sheetName, "A1", "0.00%").Success);

        using var result = ReadFormat(sheetName, "A1:B2", "stored");

        var cells = result.RootElement.GetProperty("cells");
        Assert.Equal(4, result.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(4, cells.GetArrayLength());
        Assert.Equal(["$A$1", "$B$1", "$A$2", "$B$2"],
            cells.EnumerateArray().Select(cell => cell.GetProperty("address").GetString()));
        var first = cells[0].GetProperty("stored");
        var second = cells[1].GetProperty("stored");
        Assert.True(first.GetProperty("font").GetProperty("bold").GetBoolean());
        Assert.False(second.GetProperty("font").GetProperty("bold").GetBoolean());
        Assert.Equal(18, second.GetProperty("font").GetProperty("size").GetDouble());
        Assert.Equal("#123456", first.GetProperty("font").GetProperty("color").GetProperty("rgb").GetString());
        Assert.Equal("#ABCDEF", first.GetProperty("fill").GetProperty("color").GetProperty("rgb").GetString());
        Assert.Equal("0.00%", first.GetProperty("numberFormat").GetString());
        Assert.True(first.GetProperty("wrapText").GetBoolean());
        Assert.Equal(8, first.GetProperty("borders").GetArrayLength());
    }

    [Fact]
    public void GetFormat_DisplayedIncludesConditionalFormattingWithoutChangingStoredFormat()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [[10]]).Success);
        _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1"],
            formatOptions = new
            {
                fillColor = "#FFFFFF"
            }
        });
        _fixture.Send("conditionalformat.add-rule", new
        {
            sheetName,
            rangeAddress = "A1",
            ruleType = "expression",
            formula1 = "=A1>0",
            interiorColor = "#FF0000"
        });

        using var result = ReadFormat(sheetName, "A1", "both");

        var cell = result.RootElement.GetProperty("cells")[0];
        Assert.Equal("#FFFFFF", cell.GetProperty("stored").GetProperty("fill")
            .GetProperty("color").GetProperty("rgb").GetString());
        Assert.Equal("#FF0000", cell.GetProperty("displayed").GetProperty("fill")
            .GetProperty("color").GetProperty("rgb").GetString());
    }

    [Fact]
    public void GetFormat_DoesNotSampleOrExpandRequestedScope()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);

        using var complete = ReadFormat(sheetName, "A1:A32", "stored");
        Assert.Equal(32, complete.RootElement.GetProperty("cells").GetArrayLength());
        using var single = ReadFormat(sheetName, "A1", "stored");
        Assert.Equal(1, single.RootElement.GetProperty("cells").GetArrayLength());
    }

    [Fact]
    public async Task GetFormat_InvalidViewIsNotAStoredRead()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = await _fixture.SendForFailureAsync("rangeformat.get-format",
            new { sheetName, rangeAddress = "A1", view = "unknown" });
        Assert.Contains("view", response.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetFormat_ReportsThemesProtectionAndIndividualBorders()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Font? font = null;
            Excel.Borders? borders = null;
            Excel.Border? border = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["A1"];
                font = cell.Font;
                font.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
                font.TintAndShade = 0.4;
                cell.Locked = false;
                cell.FormulaHidden = true;
                cell.IndentLevel = 2;
                borders = cell.Borders;
                border = borders[Excel.XlBordersIndex.xlDiagonalUp];
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = Excel.XlBorderWeight.xlThin;
                border.Color = 255;
            }
            finally
            {
                ComUtilities.Release(ref border);
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        using var result = ReadFormat(sheetName, "A1", "stored");
        var snapshot = result.RootElement.GetProperty("cells")[0].GetProperty("stored");
        var color = snapshot.GetProperty("font").GetProperty("color");
        Assert.Equal((int)Excel.XlThemeColor.xlThemeColorAccent2, color.GetProperty("themeColor").GetInt32());
        Assert.Equal(0.4, color.GetProperty("tintAndShade").GetDouble(), 4);
        Assert.False(snapshot.GetProperty("locked").GetBoolean());
        Assert.True(snapshot.GetProperty("formulaHidden").GetBoolean());
        Assert.Equal(2, snapshot.GetProperty("indentLevel").GetInt32());
        var diagonal = Assert.Single(snapshot.GetProperty("borders").EnumerateArray(),
            item => item.GetProperty("edge").GetString() == "xlDiagonalUp");
        Assert.Equal((int)Excel.XlLineStyle.xlContinuous, diagonal.GetProperty("lineStyle").GetInt32());
        Assert.Equal("#FF0000", diagonal.GetProperty("color").GetProperty("rgb").GetString());
        Assert.False(string.IsNullOrEmpty(snapshot.GetProperty("styleName").GetString()));
    }

    [Fact]
    public void GetFormat_RichTextDoesNotInventUniformFontProperties()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [["AB"]]).Success);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Characters? characters = null;
            Excel.Font? font = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["A1"];
                characters = cell.Characters[1, 1];
                font = characters.Font;
                font.Bold = true;
                font.Color = 255;
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref characters);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        using var result = ReadFormat(sheetName, "A1", "stored");
        var snapshot = result.RootElement.GetProperty("cells")[0].GetProperty("stored");
        Assert.Contains("font.bold", snapshot.GetProperty("mixedFields").EnumerateArray()
            .Select(field => field.GetString()));
        Assert.False(snapshot.GetProperty("font").TryGetProperty("bold", out _));
        Assert.Contains(snapshot.GetProperty("mixedFields").EnumerateArray(),
            field => field.GetString()!.StartsWith("font.color", StringComparison.Ordinal));
    }

    [Fact]
    public void GetFormat_IncludesGradientGeometryAndEveryColorStop()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Interior? interior = null;
            Excel.LinearGradient? gradient = null;
            Excel.ColorStops? stops = null;
            Excel.ColorStop? stop = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["A1"];
                interior = cell.Interior;
                interior.Pattern = Excel.XlPattern.xlPatternLinearGradient;
                gradient = (Excel.LinearGradient)interior.Gradient;
                gradient.Degree = 45;
                stops = gradient.ColorStops;
                Assert.Equal(2, stops.Count);
                stop = stops[1];
                stop.Color = 255;
            }
            finally
            {
                ComUtilities.Release(ref stop);
                ComUtilities.Release(ref stops);
                ComUtilities.Release(ref gradient);
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        using var result = ReadFormat(sheetName, "A1", "stored");
        var gradientRead = result.RootElement.GetProperty("cells")[0].GetProperty("stored")
            .GetProperty("fill").GetProperty("gradient");
        Assert.Equal("linear", gradientRead.GetProperty("kind").GetString());
        Assert.Equal(45, gradientRead.GetProperty("degree").GetDouble());
        Assert.Equal(2, gradientRead.GetProperty("stops").GetArrayLength());
        Assert.Equal("#FF0000", gradientRead.GetProperty("stops")[0].GetProperty("color")
            .GetProperty("rgb").GetString());
        Assert.Equal(0, gradientRead.GetProperty("stops")[0].GetProperty("position").GetDouble());
        Assert.Equal(1, gradientRead.GetProperty("stops")[1].GetProperty("position").GetDouble());
    }

    private JsonDocument ReadFormat(string sheetName, string rangeAddress, string view)
    {
        var response = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress, view });
        var document = JsonDocument.Parse(response.Result!);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        return document;
    }

    [Fact]
    public void GetFormat_DefaultViewIsStored()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress = "A1" });
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("stored", result.RootElement.GetProperty("view").GetString());
        Assert.True(result.RootElement.GetProperty("cells")[0].TryGetProperty("stored", out _));
        Assert.False(result.RootElement.GetProperty("cells")[0].TryGetProperty("displayed", out _));
    }

    [Fact]
    public void GetFormat_NamedAndOverlappingScopesRemainExactWithoutChangingSelection()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Format_{Guid.NewGuid():N}";
        _fixture.Send("namedrange.create", new
        {
            name,
            reference = $"'{sheetName}'!$A$1:$A$2"
        });
        _fixture.RegisterNamedRangeForCleanup(name);
        var before = ReadSelection();
        using var named = ReadFormat("", name, "displayed");
        Assert.Equal(sheetName, named.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal(2, named.RootElement.GetProperty("cellCount").GetInt64());
        Assert.False(named.RootElement.GetProperty("cells")[0].TryGetProperty("stored", out _));
        using var union = ReadFormat(sheetName, "A1:A2,A2:A3,C1", "both");
        Assert.Equal(4, union.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(["$A$1", "$C$1", "$A$2", "$A$3"], union.RootElement.GetProperty("cells")
            .EnumerateArray().Select(cell => cell.GetProperty("address").GetString()));
        Assert.Equal(before, ReadSelection());
    }

    [Fact]
    public void GetFormat_RectangularGradientRetainsGeometryAndThemeStops()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            Excel.Interior? interior = null;
            Excel.RectangularGradient? gradient = null;
            Excel.ColorStops? stops = null;
            Excel.ColorStop? stop = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cell = sheet.Range["A1"];
                interior = cell.Interior;
                interior.Pattern = Excel.XlPattern.xlPatternRectangularGradient;
                gradient = (Excel.RectangularGradient)interior.Gradient;
                gradient.RectangleTop = 0.2;
                gradient.RectangleBottom = 0.8;
                gradient.RectangleLeft = 0.1;
                gradient.RectangleRight = 0.9;
                stops = gradient.ColorStops;
                stop = stops.Add(0.4);
                stop.ThemeColor = (int)Excel.XlThemeColor.xlThemeColorAccent3;
                stop.TintAndShade = 0.25;
            }
            finally
            {
                ComUtilities.Release(ref stop);
                ComUtilities.Release(ref stops);
                ComUtilities.Release(ref gradient);
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
        using var result = ReadFormat(sheetName, "A1", "both");
        foreach (var view in new[] { "stored", "displayed" })
        {
            var gradient = result.RootElement.GetProperty("cells")[0].GetProperty(view)
                .GetProperty("fill").GetProperty("gradient");
            Assert.Equal("rectangular", gradient.GetProperty("kind").GetString());
            Assert.Equal(0.2, gradient.GetProperty("top").GetDouble(), 4);
            Assert.Equal(0.8, gradient.GetProperty("bottom").GetDouble(), 4);
            Assert.Equal(0.1, gradient.GetProperty("left").GetDouble(), 4);
            Assert.Equal(0.9, gradient.GetProperty("right").GetDouble(), 4);
            Assert.Equal(3, gradient.GetProperty("stops").GetArrayLength());
            var stop = gradient.GetProperty("stops")[1];
            Assert.Equal(0.4, stop.GetProperty("position").GetDouble(), 4);
            Assert.Equal((int)Excel.XlThemeColor.xlThemeColorAccent3,
                stop.GetProperty("color").GetProperty("themeColor").GetInt32());
            Assert.Equal(0.25, stop.GetProperty("color").GetProperty("tintAndShade").GetDouble(), 4);
        }
    }

    private (string Sheet, string Address) ReadSelection()
    {
        return _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? activeSheet = null;
            Excel.Range? selection = null;
            try
            {
                activeSheet = (Excel.Worksheet)context.App.ActiveSheet;
                selection = (Excel.Range)context.App.Selection;
                return (activeSheet.Name, selection.Address);
            }
            finally
            {
                ComUtilities.Release(ref selection);
                ComUtilities.Release(ref activeSheet);
            }
        });
    }
}
