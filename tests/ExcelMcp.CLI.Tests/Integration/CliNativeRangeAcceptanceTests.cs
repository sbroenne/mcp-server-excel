using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Sequential")]
[Trait("Layer", "CLI")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
[Trait("Acceptance", "Required")]
public sealed class CliNativeRangeAcceptanceTests(ITestOutputHelper output) : CliNativeWorkbook(output)
{
    [Fact]
    public async Task ProtectedValues_ConstantsAndSaveReopen_PreserveExactContent()
    {
        await ValuesAsync("A1", "[[424242]]");
        var rejected = await RejectedAsync("range", "set-values", "--session", Session, "--sheet-name", "Data",
            "--range-address", "A1", "--values", "[[1]]");
        Assert.Equal("Conflict", Text(rejected, "errorCategory"));
        Assert.Contains("$A$1", Text(rejected, "errorMessage"), StringComparison.Ordinal);
        Assert.Equal(424242, Number(await RangeAsync("range", "get-values", "A1"), "values.0.0"));
        await RangeAsync("range", "set-values", "A1", "--values", "[[424242]]", "--overwrite-policy", "allow");
        var constants = await RangeAsync("range", "get-special-cells", "A1:A3", "--cell-kind", "constants");
        Assert.Equal("Data", Text(constants, "sheetName"));
        Assert.Equal("$A$1:$A$3", Text(constants, "rangeAddress"));
        Assert.Equal("constants", Text(constants, "cellKind"));
        Assert.Equal(1, Number(constants, "cellCount"));
        Assert.Equal("$A$1", Assert.Single(Items(constants, "areas")).GetString());
        await ReopenAsync();
        var values = await RangeAsync("range", "get-values", "A1");
        Assert.Single(Items(values, "values"));
        Assert.Equal(424242, Number(values, "values.0.0"));
        await CommandAsync("sheet", "create", "--sheet-name", "Disposable");
        await CommandAsync("sheet", "delete", "--sheet-name", "Disposable");
        var names = Items(await CommandAsync("sheet", "list"), "worksheets").Select(sheet => Text(sheet, "name"));
        Assert.Contains("Data", names);
        Assert.DoesNotContain("Disposable", names);
    }

    [Fact]
    public async Task DynamicArray_ReportsNativeSourceAndResultBounds()
    {
        await RangeAsync("range", "set-formulas", "F1", "--formulas", """[["=SEQUENCE(3)"]]""");
        var spill = await RangeAsync("range", "get-spill-info", "F1:F3");
        Assert.Equal("supported", Text(spill, "capability"));
        Assert.Equal(3, Number(spill, "cellCount"));
        Assert.Equal("source", Text(spill, "cells.0.state"));
        Assert.Equal("result", Text(spill, "cells.2.state"));
        Assert.Equal("$F$1", Text(spill, "cells.2.sourceAddress"));
        Assert.Equal("$F$1:$F$3", Text(spill, "cells.2.spillAddress"));
    }

    [Fact]
    public async Task NativeFill_R1C1AndDependencyCoverage_RoundTrip()
    {
        await ValuesAsync("P1:P2", "[[1],[3]]");
        await CommandAsync("rangeedit", "auto-fill", "--sheet-name", "Data", "--source-range", "P1:P2",
            "--destination-range", "P1:P4", "--fill-type", "series");
        Assert.Equal(7, Number(await RangeAsync("range", "get-values", "P4"), "values.0.0"));
        await RangeAsync("range", "set-formulas", "Q1", "--formulas", """[["=RC[-1]*2"]]""", "--reference-style", "r1c1");
        await RangeAsync("rangeedit", "fill", "Q1:Q4", "--direction", "down");
        var formula = await RangeAsync("range", "get-formulas", "Q4", "--reference-style", "r1c1");
        Assert.Equal("=RC[-1]*2", Text(formula, "formulas.0.0"));
        Assert.Equal(14, Number(formula, "values.0.0"));
        var precedents = await RangeAsync("range", "trace-precedents", "Q4");
        Assert.Equal(2, Items(precedents, "nodes").Length);
        Assert.Single(Items(precedents, "edges"));
        Assert.False(Bool(precedents, "coverage.workbookComplete"));
        Assert.True(Bool(precedents, "coverage.nativeTraversalComplete"));
        Assert.Empty(Items(precedents, "unresolved"));
        var dependents = await RangeAsync("range", "trace-dependents", "P4");
        Assert.Equal(2, Items(dependents, "nodes").Length);
        Assert.Single(Items(dependents, "edges"));
        Assert.False(Bool(dependents, "coverage.workbookComplete"));
        Assert.False(Bool(dependents, "coverage.nativeTraversalComplete"));
        Assert.Single(Items(dependents, "unresolved"));
        await ValuesAsync("R1", "[[5]]");
        await RangeAsync("rangeedit", "create-series", "R1:R4", "--orientation", "columns", "--step-value", "5");
        Assert.Equal(20, Number(await RangeAsync("range", "get-values", "R4"), "values.0.0"));
    }

    [Fact]
    public async Task NativeCleanup_DuplicatesAndTextColumns_PreserveExactBoundsAndValues()
    {
        await ValuesAsync("X20:Y23", """[["Key","Amount"],[1,10],[1,20],[2,30]]""");
        var removed = await RangeAsync("rangeedit", "remove-duplicates", "X20:Y23", "--key-columns", "[1]", "--has-headers", "true");
        Assert.Equal(1, Number(removed, "removedRows"));
        Assert.Equal(2, Number(removed, "remainingRows"));
        Assert.Equal("$X$20:$Y$22", Text(removed, "remainingRange"));
        var remaining = await RangeAsync("range", "get-values", "X21:Y23");
        Assert.Equal(10, Number(remaining, "values.0.1"));
        Assert.Equal(30, Number(remaining, "values.1.1"));
        Assert.Equal(System.Text.Json.JsonValueKind.Null, At(remaining, "values.2.0").ValueKind);
        await ValuesAsync("AB20:AB21", """[["001,10,"],["002,20,"]]""");
        var split = await CommandAsync("rangeedit", "text-to-columns", "--sheet-name", "Data", "--source-range", "AB20:AB21",
            "--destination-cell", "AD20", "--options", """{"comma":true,"fields":[{"position":1,"dataType":"Text"}]}""");
        Assert.Equal(3, Number(split, "outputColumns"));
        Assert.Equal("$AD$20:$AF$21", Text(split, "destinationRange"));
        var parsed = await RangeAsync("range", "get-values", "AD20:AF21");
        Assert.Equal("001", Text(parsed, "values.0.0"));
        Assert.Equal("002", Text(parsed, "values.1.0"));
        Assert.Equal(20, Number(parsed, "values.1.1"));
        var trailing = At(parsed, "values.0.2");
        Assert.True(trailing.ValueKind == System.Text.Json.JsonValueKind.Null || trailing.GetString() == "");
    }
}
