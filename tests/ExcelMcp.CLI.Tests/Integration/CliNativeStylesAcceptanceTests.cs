using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "CellStyles")]
[Trait("Feature", "FineFormatting")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeStylesAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeThemeBordersAndCellStyle_ExposeCompleteDefinitions()
    {
        await CommandAsync("rangeformat", "format", "--sheet-name", "Data", "--range-addresses", "AA10:AB11",
            "--format-options", """{"fontThemeColor":5,"fillThemeColor":6,"indentLevel":2,"horizontalAlignment":"left","borders":[{"position":"DiagonalUp","lineStyle":"dash","color":"#123456"}]}""");
        var format = await RangeAsync("rangeformat", "get-format", "AA10");
        Assert.Equal(5, Number(format, "cells.0.stored.font.color.themeColor"));
        Assert.Equal(6, Number(format, "cells.0.stored.fill.color.themeColor"));
        Assert.Equal(2, Number(format, "cells.0.stored.indentLevel"));
        var border = Assert.Single(Items(format, "cells.0.stored.borders"), edge => Text(edge, "edge") == "xlDiagonalUp");
        Assert.Equal(-4115, Number(border, "lineStyle"));
        Assert.Equal("#123456", Text(border, "color.rgb"));
        var created = await CommandAsync("workbook", "create-cell-style", "--style-name", "WorkflowHighlight",
            "--source-sheet-name", "Data", "--source-cell-address", "AA10");
        Assert.Equal("WorkflowHighlight", Text(created, "style.name"));
        Assert.False(Bool(created, "style.builtIn"));
        await RangeAsync("rangeformat", "set-style", "AD10", "--style-name", "WorkflowHighlight");
        var updated = await CommandAsync("workbook", "update-cell-style", "--style-name", "WorkflowHighlight",
            "--style-options", """{"includeFont":true,"formatOptions":{"bold":true}}""");
        Assert.True(Bool(updated, "style.format.font.bold"));
        var user = await RangeAsync("rangeformat", "get-style", "AD10");
        Assert.Equal("WorkflowHighlight", Text(user, "styleName"));
        Assert.False(Bool(user, "isBuiltInStyle"));
        var definition = await CommandAsync("workbook", "get-cell-style", "--style-name", "WorkflowHighlight");
        Assert.True(Bool(definition, "style.format.font.bold"));
        Assert.Equal(6, Items(definition, "style.format.borders").Length);
        await CommandAsync("workbook", "delete-cell-style", "--style-name", "WorkflowHighlight");
    }

    [Fact]
    [Trait("Feature", "TableStyles")]
    public async Task NativeTableStyle_ClonesUpdatesListsAndDeletesAllElements()
    {
        var created = await CommandAsync("workbook", "create-table-style", "--style-name", "WorkflowTableStyle",
            "--source-style-name", "TableStyleMedium2");
        Assert.False(Bool(created, "style.builtIn"));
        Assert.Equal(43, Items(created, "style.elements").Length);
        var updated = await CommandAsync("workbook", "update-table-style", "--style-name", "WorkflowTableStyle",
            "--table-style-options", """{"elements":[{"elementType":"xlHeaderRow","fillColor":"#123456","bold":false}]}""");
        var header = Assert.Single(Items(updated, "style.elements"), element => Text(element, "elementType") == "xlHeaderRow");
        Assert.Equal("#123456", Text(header, "fill.color.rgb"));
        Assert.Equal(43, Items(await CommandAsync("workbook", "get-table-style", "--style-name", "WorkflowTableStyle"), "style.elements").Length);
        var listed = Assert.Single(Items(await CommandAsync("workbook", "list-table-styles"), "styles"),
            style => Text(style, "name") == "WorkflowTableStyle");
        Assert.False(Bool(listed, "builtIn"));
        await CommandAsync("workbook", "delete-table-style", "--style-name", "WorkflowTableStyle");
    }
}
