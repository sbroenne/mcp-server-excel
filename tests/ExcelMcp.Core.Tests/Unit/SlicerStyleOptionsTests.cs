using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Slicer;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("RequiresExcel", "false")]
[Trait("Category", "Unit")]
[Trait("Feature", "Slicer")]
[Trait("Layer", "Core")]
[Trait("Speed", "Fast")]
public sealed class SlicerStyleOptionsTests
{
    [Theory]
    [InlineData("wholeTable", Excel.XlTableStyleElementType.xlWholeTable)]
    [InlineData("header", Excel.XlTableStyleElementType.xlHeaderRow)]
    [InlineData("selectedItemWithData", Excel.XlTableStyleElementType.xlSlicerSelectedItemWithData)]
    [InlineData("selectedItemWithNoData", Excel.XlTableStyleElementType.xlSlicerSelectedItemWithNoData)]
    [InlineData("unselectedItemWithData", Excel.XlTableStyleElementType.xlSlicerUnselectedItemWithData)]
    [InlineData("unselectedItemWithNoData", Excel.XlTableStyleElementType.xlSlicerUnselectedItemWithNoData)]
    [InlineData("hoveredSelectedItemWithData", Excel.XlTableStyleElementType.xlSlicerHoveredSelectedItemWithData)]
    [InlineData("hoveredSelectedItemWithNoData", Excel.XlTableStyleElementType.xlSlicerHoveredSelectedItemWithNoData)]
    [InlineData("hoveredUnselectedItemWithData", Excel.XlTableStyleElementType.xlSlicerHoveredUnselectedItemWithData)]
    [InlineData("hoveredUnselectedItemWithNoData", Excel.XlTableStyleElementType.xlSlicerHoveredUnselectedItemWithNoData)]
    public void Mapping_UsesNativeElementsAndKeepsExplicitFalse(string property, Excel.XlTableStyleElementType native)
    {
        var options = Deserialize($$$"""{"{{{property}}}":{"fillColor":"#123456","fontColor":"#234567","borderColor":"#345678","bold":false}}""");
        var mapped = options.ToTableStyleOptions();
        WorkbookCommands.ValidateTableStyleOptions(mapped);
        Assert.True(mapped.ShowAsAvailableSlicerStyle);
        var element = Assert.Single(mapped.Elements!);
        Assert.Equal(native.ToString(), element.ElementType);
        Assert.Equal("#123456", element.FillColor);
        Assert.Equal("#234567", element.FontColor);
        Assert.False(element.Bold);
        Assert.Equal(4, element.Borders!.Count);
        Assert.Equal(4, element.Borders.Select(border => border.Position).Distinct().Count());
        Assert.All(element.Borders, border =>
        {
            Assert.Equal("#345678", border.Color);
            Assert.Equal("continuous", border.LineStyle);
            Assert.Equal("thin", border.Weight);
        });
    }

    [Fact]
    public void OmittedSettings_DoNotInventFormatting()
    {
        Assert.Empty(Deserialize("{}").ToTableStyleOptions().Elements!);
        var format = Assert.Single(Deserialize("""{"header":{"bold":true}}""").ToTableStyleOptions().Elements!);
        Assert.True(format.Bold);
        Assert.Null(format.FillColor);
        Assert.Null(format.FontColor);
        Assert.Null(format.Borders);
    }

    [Theory]
    [InlineData("""{"unknown":{}}""")]
    [InlineData("""{"header":{"unknown":true}}""")]
    public void UnknownProperties_Fail(string json) =>
        Assert.Throws<JsonException>(() => Deserialize(json));

    [Theory]
    [InlineData("fillColor")]
    [InlineData("fontColor")]
    [InlineData("borderColor")]
    public void InvalidColours_FailBeforeExcel(string property)
    {
        var options = Deserialize($$$"""{"header":{"{{{property}}}":"invalid"}}""").ToTableStyleOptions();
        Assert.Throws<ArgumentException>(() => WorkbookCommands.ValidateTableStyleOptions(options));
    }

    private static SlicerStyleOptions Deserialize(string json) =>
        JsonSerializer.Deserialize<SlicerStyleOptions>(json, JsonSerializerOptions.Web)!;
}
