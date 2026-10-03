using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "FineFormatting")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceFineFormattingTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public async Task Format_SharedTypedOptionsPersistAndLeaveGapsUnchanged()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string gap = WithoutBorders(Read(sheetName, "C1"));
        var response = _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1:B2", "D1:E2"],
            formatOptions = new
            {
                bold = true,
                fontThemeColor = 5,
                fontTintAndShade = 0.25,
                fillThemeColor = 6,
                fillTintAndShade = -0.2,
                underline = "doubleAccounting",
                strikethrough = true,
                horizontalAlignment = "left",
                indentLevel = 2,
                shrinkToFit = true,
                readingOrder = "rightToLeft",
                numberFormat = "0.00 \"Total\"",
                borders = new[]
                {
                    new { position = "Left", lineStyle = "double", weight = "thin", themeColor = 7, tintAndShade = 0.1 },
                    new { position = "InsideHorizontal", lineStyle = "continuous", weight = "thin", themeColor = 8, tintAndShade = 0.0 },
                    new { position = "DiagonalUp", lineStyle = "dash", weight = "thin", themeColor = 9, tintAndShade = 0.0 }
                }
            }
        });
        Assert.True(response.Success, response.ErrorMessage);
        string beforeSave = Read(sheetName, "A1");
        Assert.Equal(gap, WithoutBorders(Read(sheetName, "C1")));
        Assert.Equal(beforeSave, Read(sheetName, "D1"));
        using (var read = JsonDocument.Parse(beforeSave))
        {
            var cell = read.RootElement;
            Assert.True(cell.GetProperty("font").GetProperty("bold").GetBoolean());
            Assert.Equal(5, cell.GetProperty("font").GetProperty("color").GetProperty("themeColor").GetInt32());
            Assert.Equal(6, cell.GetProperty("fill").GetProperty("color").GetProperty("themeColor").GetInt32());
            Assert.Equal(0.25, cell.GetProperty("font").GetProperty("color").GetProperty("tintAndShade").GetDouble(), 3);
            Assert.Equal(-0.2, cell.GetProperty("fill").GetProperty("color").GetProperty("tintAndShade").GetDouble(), 3);
            Assert.Equal(2, cell.GetProperty("indentLevel").GetInt32());
            Assert.Equal(5, cell.GetProperty("font").GetProperty("underline").GetInt32());
            Assert.True(cell.GetProperty("font").GetProperty("strikethrough").GetBoolean());
            Assert.True(cell.GetProperty("shrinkToFit").GetBoolean());
            Assert.Equal(-5004, cell.GetProperty("readingOrder").GetInt32());
            Assert.Equal("0.00 \"Total\"", cell.GetProperty("numberFormat").GetString());
            Assert.Contains(cell.GetProperty("borders").EnumerateArray(),
                border => border.GetProperty("edge").GetString() == "xlDiagonalUp" &&
                          border.GetProperty("lineStyle").GetInt32() == -4115);
        }
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Borders? borders = null;
            Excel.Border? border = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                foreach (var (address, edge, style, weight, theme, tint) in new[]
                {
                    // Excel promotes a double border to thick even when thin was requested.
                    ("A1:B2", Excel.XlBordersIndex.xlEdgeLeft, -4119, 4, 7, 0.1),
                    ("A1:B2", Excel.XlBordersIndex.xlInsideHorizontal, 1, 2, 8, 0.0),
                    // Aggregate diagonal color is DBNull; inspect each cell instead.
                    ("A1", Excel.XlBordersIndex.xlDiagonalUp, -4115, 2, 9, 0.0),
                    ("B1", Excel.XlBordersIndex.xlDiagonalUp, -4115, 2, 9, 0.0),
                    ("A2", Excel.XlBordersIndex.xlDiagonalUp, -4115, 2, 9, 0.0),
                    ("B2", Excel.XlBordersIndex.xlDiagonalUp, -4115, 2, 9, 0.0)
                })
                {
                    range = sheet!.Range[address];
                    borders = range.Borders;
                    border = borders[edge];
                    Assert.Equal(style, Convert.ToInt32(border.LineStyle, CultureInfo.InvariantCulture));
                    Assert.Equal(weight, Convert.ToInt32(border.Weight, CultureInfo.InvariantCulture));
                    Assert.Equal(theme, Convert.ToInt32(border.ThemeColor, CultureInfo.InvariantCulture));
                    Assert.Equal(tint, border.TintAndShade, 3);
                    ComUtilities.Release(ref border);
                    ComUtilities.Release(ref borders);
                    ComUtilities.Release(ref range);
                }
            }
            finally
            {
                ComUtilities.Release(ref border);
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
        await _fixture.SaveAndReopenAsync();
        Assert.Equal(beforeSave, Read(sheetName, "A1"));
        Assert.Equal(beforeSave, Read(sheetName, "D1"));
    }

    [Fact]
    public async Task Format_InvalidLaterTargetPreservesEarlierTarget()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string before = Read(sheetName, "A1");
        var response = await _fixture.SendForFailureAsync("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1", "NotARange"],
            formatOptions = new { bold = true, fillColor = "#FF0000" }
        });
        Assert.False(response.Success);
        Assert.Contains("index 1", response.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(before, Read(sheetName, "A1"));
    }

    [Theory]
    [InlineData("Left", 7)]
    [InlineData("Top", 8)]
    [InlineData("Bottom", 9)]
    [InlineData("Right", 10)]
    [InlineData("InsideVertical", 11)]
    [InlineData("InsideHorizontal", 12)]
    [InlineData("DiagonalDown", 5)]
    [InlineData("DiagonalUp", 6)]
    public void Format_EachNativeBorderCanBeSetAndRemoved(string position, int index)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1:C3"],
            formatOptions = new { borders = new[] { new { position, lineStyle = "dash", weight = "thin", color = "#123456" } } }
        });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.Equal((-4115, 2, 0x563412), ReadBorder(sheetName, index));
        response = _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1:C3"],
            formatOptions = new { borders = new[] { new { position, lineStyle = "none" } } }
        });
        Assert.True(response.Success, response.ErrorMessage);
        Assert.Equal(-4142, ReadBorder(sheetName, index).Style);
    }

    [Theory]
    [InlineData("none", -4142)]
    [InlineData("single", 2)]
    [InlineData("double", -4119)]
    [InlineData("singleAccounting", 4)]
    [InlineData("doubleAccounting", 5)]
    public void Format_UnderlineUsesNativeKinds(string underline, int expected)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1"],
            formatOptions = new { underline }
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var read = JsonDocument.Parse(Read(sheetName, "A1"));
        Assert.Equal(expected, read.RootElement.GetProperty("font").GetProperty("underline").GetInt32());
    }

    [Theory]
    [InlineData("left", -4131)]
    [InlineData("center", -4108)]
    [InlineData("right", -4152)]
    [InlineData("justify", -4130)]
    [InlineData("distributed", -4117)]
    [InlineData("fill", 5)]
    [InlineData("centerAcrossSelection", 7)]
    public void Format_HorizontalAlignmentsUseNativeValues(string horizontalAlignment, int expected)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1"],
            formatOptions = new { horizontalAlignment }
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var read = JsonDocument.Parse(Read(sheetName, "A1"));
        Assert.Equal(expected, read.RootElement.GetProperty("horizontalAlignment").GetInt32());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Format_ThemeFontsReadBackFromExcel(int themeFont)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        var response = _fixture.Send("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1"],
            formatOptions = new { themeFont }
        });
        Assert.True(response.Success, response.ErrorMessage);
        using var read = JsonDocument.Parse(Read(sheetName, "A1"));
        Assert.Equal(themeFont, read.RootElement.GetProperty("font").GetProperty("themeFont").GetInt32());
    }

    [Theory]
    [InlineData("""{"fontColor":"#123456","fontThemeColor":5}""")]
    [InlineData("""{"fillColor":"#123456","fillThemeColor":5}""")]
    [InlineData("""{"fontThemeColor":0}""")]
    [InlineData("""{"fillThemeColor":13}""")]
    [InlineData("""{"fontTintAndShade":1.1}""")]
    [InlineData("""{"fillTintAndShade":-1.1}""")]
    [InlineData("""{"themeFont":3}""")]
    [InlineData("""{"fontName":"Arial","themeFont":1}""")]
    [InlineData("""{"fontSize":0}""")]
    [InlineData("""{"fontSize":410}""")]
    [InlineData("""{"indentLevel":-1}""")]
    [InlineData("""{"indentLevel":16}""")]
    [InlineData("""{"orientation":91}""")]
    [InlineData("""{"subscript":true,"superscript":true}""")]
    [InlineData("""{"underline":"bad"}""")]
    [InlineData("""{"readingOrder":"bad"}""")]
    [InlineData("""{"unknown":true}""")]
    [InlineData("""{"borders":[{"lineStyle":"dash"}]}""")]
    [InlineData("""{"borders":[{"position":"bad","lineStyle":"dash"}]}""")]
    [InlineData("""{"borders":[{"position":"Left"}]}""")]
    [InlineData("""{"borders":[{"position":"Left","lineStyle":"none","weight":"thin"}]}""")]
    [InlineData("""{"borders":[{"position":"Left","lineStyle":"bad"}]}""")]
    [InlineData("""{"borders":[{"position":"Left","color":"#123456","themeColor":5}]}""")]
    [InlineData("""{"borders":[{"position":"Left","lineStyle":"dash"},{"position":"Left","lineStyle":"double"}]}""")]
    public async Task Format_InvalidOptionsPreserveAllTargets(string json)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        string before = Read(sheetName, "A1");
        using var options = JsonDocument.Parse(json);
        var response = await _fixture.SendForFailureAsync("rangeformat.format", new
        {
            sheetName,
            rangeAddresses = (string[])["A1", "C1"],
            formatOptions = options.RootElement
        });
        Assert.False(response.Success);
        Assert.False(string.IsNullOrWhiteSpace(response.ErrorMessage));
        Assert.Equal(before, Read(sheetName, "A1"));
        Assert.Equal(before, Read(sheetName, "C1"));
    }

    [Theory]
    [InlineData("format-range")]
    [InlineData("format-ranges")]
    public async Task Format_ObsoleteActionsAreRemoved(string action)
    {
        var response = await _fixture.SendForFailureAsync($"rangeformat.{action}", new { });
        Assert.False(response.Success);
        Assert.Contains("Unknown action", response.ErrorMessage, StringComparison.Ordinal);
    }

    private (int Style, int Weight, int Color) ReadBorder(string sheetName, int index) =>
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Borders? borders = null;
            Excel.Border? border = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range[index is 5 or 6 ? "A1" : "A1:C3"];
                borders = range.Borders;
                border = borders[(Excel.XlBordersIndex)index];
                return (Convert.ToInt32((object)border.LineStyle, CultureInfo.InvariantCulture),
                    Convert.ToInt32((object)border.Weight, CultureInfo.InvariantCulture),
                    Convert.ToInt32((object)border.Color, CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref border);
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private string Read(string sheetName, string rangeAddress)
    {
        var response = _fixture.Send("rangeformat.get-format", new { sheetName, rangeAddress });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        return result.RootElement.GetProperty("cells")[0].GetProperty("stored").GetRawText();
    }

    private static string WithoutBorders(string snapshot)
    {
        // Borders and row heights are shared with adjacent cells; exclude those native layout effects.
        var cell = JsonNode.Parse(snapshot)!.AsObject();
        Assert.True(cell.Remove("borders"));
        Assert.True(cell.Remove("rowHeight"));
        var fields = cell["mixedFields"]!.AsArray();
        for (int index = fields.Count - 1; index >= 0; index--)
            if (fields[index]!.GetValue<string>().StartsWith("borders.", StringComparison.Ordinal))
                fields.RemoveAt(index);
        return cell.ToJsonString();
    }
}
