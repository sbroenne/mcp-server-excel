using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativePivotAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeCalculationLayoutFiltersAndSource_SurviveReopen()
    {
        await CreatePivotAsync();
        var calculation = await PivotAsync("pivottablefield", "set-field-calculation", "--field-name", "Total Sales",
            "--calculation", "PercentOfTotal");
        Assert.Equal("PercentOfTotal", Text(calculation, "calculation"));
        Assert.Equal("Sum", Text(calculation, "function"));
        var data = await PivotAsync("pivottablecalc", "get-data");
        Assert.Equal(0.25, Number(data, "values.1.1"));
        Assert.Equal(0.75, Number(data, "values.2.1"));
        var layout = await PivotAsync("pivottablecalc", "set-layout-options",
            "--layout-options", """{"rowLayout":1,"repeatLabels":true,"styleName":"PivotStyleMedium9","preserveFormatting":true}""");
        Assert.Equal("PivotStyleMedium9", Text(layout, "styleName"));
        Assert.True(Bool(layout, "rowFields.0.repeatLabels"));
        var added = await PivotAsync("pivottablefield", "add-field-filter", "--field-name", "Region",
            "--filter-options", """{"type":"CaptionEquals","text1":"North"}""");
        Assert.Equal("North", Text(added, "filters.0.value1"));
        var filters = await PivotAsync("pivottablefield", "get-field-filters", "--field-name", "Region");
        Assert.Equal("CaptionEquals", Text(Assert.Single(Items(filters, "filters")), "type"));
        Assert.Empty(Items(await PivotAsync("pivottablefield", "clear-field-filters", "--field-name", "Region"), "filters"));
        var source = await PivotAsync("pivottable", "get-source");
        Assert.Equal(2, Number(source, "recordCount"));
        Assert.Equal("CalculationPivot", Text(source, "sharedPivotTables.0"));
        var replaced = await PivotAsync("pivottable", "set-source", "--source-sheet-name", "Data", "--source-range-address", "U1:V3");
        Assert.Equal(2, Number(replaced, "recordCount"));
        Assert.Empty(Items(replaced, "connectedSlicerCaches"));
        await ReopenAsync();
        var saved = Assert.Single(Items(await PivotAsync("pivottablefield", "list-fields"), "valueFields"));
        Assert.Equal("Total Sales", Text(saved, "fieldName"));
        Assert.Equal("PercentOfTotal", Text(saved, "calculation"));
        Assert.Equal("Sum", Text(saved, "function"));
    }
}

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "Slicer")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeSlicerAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeReportControl_UpdatesLayoutWithoutChangingPivotCalculation()
    {
        await CreatePivotAsync();
        await PivotAsync("pivottablefield", "set-field-calculation", "--field-name", "Total Sales", "--calculation", "PercentOfTotal");
        await CommandAsync("slicer", "create-slicer", "--pivot-table-name", "CalculationPivot", "--field-name", "Region",
            "--slicer-name", "ReportRegions", "--destination-sheet", "Data", "--position", "AB1");
        var updated = await CommandAsync("slicer", "update-slicer", "--slicer-name", "ReportRegions",
            "--slicer-options", """{"width":240,"height":180,"columnCount":2,"caption":"Regions","displayHeader":false}""");
        Assert.Equal(240, Number(updated, "slicer.width"));
        Assert.Equal(2, Number(updated, "slicer.columnCount"));
        var complete = await CommandAsync("slicer", "get-slicer", "--slicer-name", "ReportRegions");
        Assert.Equal("Regions", Text(complete, "slicer.caption"));
        Assert.Equal(2, Items(complete, "slicer.availableItems").Length);
        Assert.Equal("CalculationPivot", Text(complete, "slicer.connectedPivotTables.0"));
        await CommandAsync("slicer", "delete-slicer", "--slicer-name", "ReportRegions");
        var unchanged = await PivotAsync("pivottablecalc", "get-data");
        Assert.Equal(0.25, Number(unchanged, "values.1.1"));
        Assert.Equal(0.75, Number(unchanged, "values.2.1"));
    }
}
