using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "StructuredFilters")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeFilterAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeTwoConditionFilter_ReportsBothCriteriaAndMatchingRows()
    {
        await CommandAsync("sheet", "create", "--sheet-name", "Filters");
        await CommandAsync("range", "set-values", "--sheet-name", "Filters", "--range-address", "A1:B6",
            "--values", """[["Category","Amount"],["A",10],["B",20],["C",30],["A",40],["B",50]]""");
        await CommandAsync("rangeedit", "apply-filter", "--sheet-name", "Filters", "--range-address", "A1:B6", "--column-index", "2",
            "--filter-options", """{"filterOperator":"And","criteria1":">=20","criteria2":"<=40"}""");
        var criteria = await CommandAsync("rangeedit", "get-filters", "--sheet-name", "Filters", "--range-address", "A1:B6");
        Assert.Equal(2, Items(criteria, "columnFilters").Length);
        Assert.Equal("And", Text(criteria, "columnFilters.1.filterOperator"));
        Assert.Equal(">=20", Text(criteria, "columnFilters.1.criteria1.value"));
        Assert.Equal("<=40", Text(criteria, "columnFilters.1.criteria2.value"));
        var visible = await CommandAsync("range", "get-special-cells", "--sheet-name", "Filters", "--range-address", "A2:A6",
            "--cell-kind", "Visible");
        Assert.Equal(3, Number(visible, "cellCount"));
        await CommandAsync("rangeedit", "clear-filters", "--sheet-name", "Filters", "--range-address", "A1:B6");
    }
}

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "PageLayout")]
[Trait("Feature", "FineFormatting")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeReportAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task NativeVisibilityPageSetupAndBreaks_PreserveExactPositions()
    {
        await RangeAsync("rangeformat", "set-visibility", "A40,A42", "--axis", "rows", "--hidden", "true");
        var visibility = await RangeAsync("rangeformat", "get-visibility", "A40:A42", "--axis", "rows");
        Assert.Equal(3, Items(visibility, "items").Length);
        Assert.True(Bool(visibility, "items.0.hidden"));
        Assert.False(Bool(visibility, "items.1.hidden"));
        Assert.True(Bool(visibility, "items.2.hidden"));
        Assert.Equal("undetermined", Text(visibility, "items.0.hiddenCause"));
        await RangeAsync("rangeformat", "set-visibility", "A40,A42", "--axis", "rows", "--hidden", "false");
        await CommandAsync("sheet", "create", "--sheet-name", "Filters");
        await CommandAsync("range", "set-values", "--sheet-name", "Filters", "--range-address", "A1:B6",
            "--values", """[["Category","Amount"],["A",10],["B",20],["C",30],["A",40],["B",50]]""");
        await CommandAsync("worksheetstyle", "set-page-setup", "--sheet-name", "Filters",
            "--page-setup-options", """{"printArea":"A1:B6","leftMargin":36,"centerHeader":"Report","zoomPercent":100}""");
        var setup = await CommandAsync("worksheetstyle", "get-page-setup", "--sheet-name", "Filters");
        Assert.Equal("$A$1:$B$6", Text(setup, "printArea"));
        Assert.Equal(36, Number(setup, "leftMargin"));
        Assert.Equal("Report", Text(setup, "centerHeader"));
        await CommandAsync("worksheetstyle", "set-page-breaks", "--sheet-name", "Filters",
            "--page-break-options", """{"rows":[4],"columns":[]}""");
        var breaks = await CommandAsync("worksheetstyle", "get-page-breaks", "--sheet-name", "Filters");
        Assert.Single(Items(breaks, "horizontal"), item => Bool(item, "isManual") && Number(item, "position") == 4);
    }
}
