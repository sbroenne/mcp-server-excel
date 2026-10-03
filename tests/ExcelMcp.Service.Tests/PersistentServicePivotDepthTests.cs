using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotDepth")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePivotDepthTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPersistentPivotTableCommands _pivot =
        ServiceCommandProxy.Create<IPersistentPivotTableCommands>(fixture);

    [Fact]
    public void NativeLabelAndValueFilters_ChangeActualRowsAndCanBeCleared()
    {
        var (_, name) = CreatePivot();
        using (var added = Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "Region",
            filterOptions = new { type = "CaptionEquals", text1 = "North" }
        }))
            Assert.Single(added.RootElement.GetProperty("filters").EnumerateArray());
        AssertValues(name, 100d, ("North", 100d), ("A", 40d), ("B", 60d));
        using (Send("pivottablefield.clear-field-filters", new { pivotTableName = name, fieldName = "Region" })) { }
        using (var added = Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "Region",
            filterOptions = new { type = "ValueIsGreaterThan", number1 = 150d, dataFieldName = "Total Sales" }
        }))
            Assert.Equal("ValueIsGreaterThan", added.RootElement.GetProperty("filters")[0].GetProperty("type").GetString());
        AssertValues(name, 300d, ("South", 300d), ("A", 120d), ("B", 180d));
    }

    [Fact]
    public void LayoutOptions_ReadActualStyleAndRepeatedRowLabels()
    {
        var (_, name) = CreatePivot();
        using var set = Send("pivottablecalc.set-layout-options", new
        {
            pivotTableName = name,
            layoutOptions = new { rowLayout = 1, repeatLabels = true, styleName = "PivotStyleMedium9", preserveFormatting = true, showRowStripes = true }
        });
        Assert.Equal("PivotStyleMedium9", set.RootElement.GetProperty("styleName").GetString());
        Assert.True(set.RootElement.GetProperty("preserveFormatting").GetBoolean());
        Assert.True(set.RootElement.GetProperty("showRowStripes").GetBoolean());
        Assert.All(set.RootElement.GetProperty("rowFields").EnumerateArray(),
            field => Assert.True(field.GetProperty("repeatLabels").GetBoolean()));
    }

    [Fact]
    public void ItemExpansion_ChangesOnlyNamedParent()
    {
        var (_, name) = CreatePivot();
        using var otherBefore = Send("pivottablefield.get-item-expansion", new
        {
            pivotTableName = name,
            fieldName = "Region",
            itemName = "South"
        });
        using var collapsed = Send("pivottablefield.set-item-expansion", new
        {
            pivotTableName = name,
            fieldName = "Region",
            itemName = "North",
            expanded = false
        });
        Assert.False(collapsed.RootElement.GetProperty("expanded").GetBoolean());
        AssertValues(name, 400d, ("North", 100d), ("South", 300d), ("A", 120d), ("B", 180d));
        using var otherAfter = Send("pivottablefield.get-item-expansion", new
        {
            pivotTableName = name,
            fieldName = "Region",
            itemName = "South"
        });
        Assert.Equal(otherBefore.RootElement.GetRawText(), otherAfter.RootElement.GetRawText());
        using var expanded = Send("pivottablefield.set-item-expansion", new
        {
            pivotTableName = name,
            fieldName = "Region",
            itemName = "North",
            expanded = true
        });
        Assert.True(expanded.RootElement.GetProperty("expanded").GetBoolean());
        AssertValues(name, 400d, ("North", 100d), ("A", 40d), ("B", 60d),
            ("South", 300d), ("A", 120d), ("B", 180d));
    }

    [Fact]
    public void SourceReplacement_PreservesPlacedFieldsAndReadsNewRecords()
    {
        var (sheet, name) = CreatePivot();
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A8:C10",
            [["Region", "Product", "Sales"], ["North", "A", 250], ["South", "B", 450]]).Success);
        using var replaced = Send("pivottable.set-source", new
        {
            pivotTableName = name,
            sourceSheetName = sheet,
            sourceRangeAddress = "A8:C10"
        });
        Assert.Equal(2, replaced.RootElement.GetProperty("recordCount").GetInt32());
        AssertValues(name, 700d, ("North", 250d), ("A", 250d), ("South", 450d), ("B", 450d));
    }

    [Fact]
    public async Task SharedCacheOptions_RejectMutationWithoutChangingOtherPivot()
    {
        var (sheet, name) = CreatePivot();
        var other = CreateSharedPivot(sheet, name);
        using var targetBefore = Send("pivottable.get-cache-options", new { pivotTableName = name });
        using var original = Send("pivottable.get-cache-options", new { pivotTableName = other });
        bool refreshOnFileOpen = original.RootElement.GetProperty("refreshOnFileOpen").GetBoolean();
        var failure = await _fixture.SendForFailureAsync("pivottable.set-cache-options", new
        {
            pivotTableName = name,
            refreshOnFileOpen = !refreshOnFileOpen
        });
        Assert.False(failure.Success);
        Assert.Contains("shared", failure.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        using var unchanged = Send("pivottable.get-cache-options", new { pivotTableName = other });
        Assert.Equal(refreshOnFileOpen, unchanged.RootElement.GetProperty("refreshOnFileOpen").GetBoolean());
        Assert.Equal(original.RootElement.GetRawText(), unchanged.RootElement.GetRawText());
        using var targetAfter = Send("pivottable.get-cache-options", new { pivotTableName = name });
        Assert.Equal(targetBefore.RootElement.GetRawText(), targetAfter.RootElement.GetRawText());
    }

    [Theory]
    [InlineData("CaptionEquals", "North", null, 100d)]
    [InlineData("CaptionDoesNotEqual", "North", null, 300d)]
    [InlineData("CaptionBeginsWith", "N", null, 100d)]
    [InlineData("CaptionDoesNotBeginWith", "N", null, 300d)]
    [InlineData("CaptionEndsWith", "h", null, 400d)]
    [InlineData("CaptionDoesNotEndWith", "h", null, 0d)]
    [InlineData("CaptionContains", "ort", null, 100d)]
    [InlineData("CaptionDoesNotContain", "ort", null, 300d)]
    [InlineData("CaptionIsGreaterThan", "North", null, 300d)]
    [InlineData("CaptionIsGreaterThanOrEqualTo", "North", null, 400d)]
    [InlineData("CaptionIsLessThan", "South", null, 100d)]
    [InlineData("CaptionIsLessThanOrEqualTo", "South", null, 400d)]
    [InlineData("CaptionIsBetween", "North", "South", 400d)]
    [InlineData("CaptionIsNotBetween", "North", "South", 0d)]
    public void LabelFilters_ReturnNativeCriteriaAndActualTotals(string type, string first, string? second, double expected)
    {
        var (_, name) = CreatePivot();
        using var result = Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "Region",
            filterOptions = new { type, text1 = first, text2 = second }
        });
        var filter = Assert.Single(result.RootElement.GetProperty("filters").EnumerateArray());
        Assert.Equal(type, filter.GetProperty("type").GetString());
        Assert.Equal(first, filter.GetProperty("value1").GetString());
        if (second is not null)
            Assert.Equal(second, filter.GetProperty("value2").GetString());
        AssertTotal(name, expected);
    }

    [Theory]
    [InlineData("ValueEquals", 100, null, 100d)]
    [InlineData("ValueDoesNotEqual", 100, null, 300d)]
    [InlineData("ValueIsGreaterThan", 100, null, 300d)]
    [InlineData("ValueIsGreaterThanOrEqualTo", 100, null, 400d)]
    [InlineData("ValueIsLessThan", 300, null, 100d)]
    [InlineData("ValueIsLessThanOrEqualTo", 300, null, 400d)]
    [InlineData("ValueIsBetween", 100, 300d, 400d)]
    [InlineData("ValueIsNotBetween", 100, 300d, 0d)]
    [InlineData("TopCount", 1, null, 300d)]
    [InlineData("BottomCount", 1, null, 100d)]
    [InlineData("TopPercent", 50, null, 300d)]
    [InlineData("BottomPercent", 25, null, 100d)]
    [InlineData("TopSum", 250, null, 300d)]
    [InlineData("BottomSum", 50, null, 100d)]
    public void ValueFilters_ReturnNativeCriteriaAndActualTotals(string type, double first, double? second, double expected)
    {
        var (_, name) = CreatePivot();
        using var result = Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "Region",
            filterOptions = new { type, number1 = first, number2 = second, dataFieldName = "Total Sales" }
        });
        var filter = Assert.Single(result.RootElement.GetProperty("filters").EnumerateArray());
        Assert.Equal(type, filter.GetProperty("type").GetString());
        Assert.Equal(first, filter.GetProperty("value1").GetDouble());
        Assert.Equal("Total Sales", filter.GetProperty("dataFieldName").GetString());
        AssertTotal(name, expected);
    }

    [Theory]
    [InlineData("SpecificDate", false, 20d)]
    [InlineData("NotSpecificDate", false, 40d)]
    [InlineData("Before", false, 10d)]
    [InlineData("BeforeOrEqualTo", false, 30d)]
    [InlineData("After", false, 30d)]
    [InlineData("AfterOrEqualTo", false, 50d)]
    [InlineData("DateBetween", true, 50d)]
    [InlineData("DateNotBetween", true, 10d)]
    public void DateFilters_UseNativeDateFieldAndActualTotals(string type, bool interval, double expected)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Dates_{Guid.NewGuid():N}";
        using (Send("range.set-values", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B4",
            values = new object[][] { ["Date", "Sales"], [new DateTime(2024, 1, 10).ToOADate(), 10],
                [new DateTime(2024, 2, 10).ToOADate(), 20], [new DateTime(2024, 3, 10).ToOADate(), 30] }
        })) { }
        using (Send("range.set-number-format", new { sheetName = sheet, rangeAddress = "A2:A4", formatCode = "yyyy-mm-dd" })) { }
        var created = _pivot.CreateFromRange(_fixture.BatchToken, sheet, "A1:B4", sheet, "E1", name);
        Assert.True(created.Success, created.ErrorMessage);
        var row = _pivot.AddRowField(_fixture.BatchToken, name, "Date");
        Assert.True(row.Success, row.ErrorMessage);
        var value = _pivot.AddValueField(_fixture.BatchToken, name, "Sales", AggregationFunction.Sum, "Total Sales");
        Assert.True(value.Success, value.ErrorMessage);
        using var result = Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "Date",
            filterOptions = new { type, date1 = new DateTime(2024, 2, 10), date2 = interval ? new DateTime(2024, 3, 10) : (DateTime?)null }
        });
        Assert.Equal(type, Assert.Single(result.RootElement.GetProperty("filters").EnumerateArray()).GetProperty("type").GetString());
        AssertTotal(name, expected);
    }

    [Theory]
    [InlineData("""{"type":"CaptionEquals"}""")]
    [InlineData("""{"type":"CaptionEquals","text1":"North","number1":1}""")]
    [InlineData("""{"type":"CaptionEquals","text1":"North","text2":"South"}""")]
    [InlineData("""{"type":"CaptionIsBetween","text1":"North"}""")]
    [InlineData("""{"type":"ValueEquals","number1":100}""")]
    [InlineData("""{"type":"ValueEquals","number1":100,"dataFieldName":"Sales"}""")]
    [InlineData("""{"type":"ValueIsBetween","number1":300,"number2":100,"dataFieldName":"Total Sales"}""")]
    [InlineData("""{"type":"TopCount","number1":1.5,"dataFieldName":"Total Sales"}""")]
    [InlineData("""{"type":"TopCount","number1":0,"dataFieldName":"Total Sales"}""")]
    [InlineData("""{"type":"TopPercent","number1":101,"dataFieldName":"Total Sales"}""")]
    [InlineData("""{"type":"SpecificDate","date1":"2024-02-10"}""")]
    [InlineData("""{"type":"DateBetween","date1":"2024-02-10"}""")]
    [InlineData("""{"type":"Invalid"}""")]
    [InlineData("""{"type":"CaptionEquals","text1":"North","unknown":true}""")]
    public async Task InvalidFilter_PreservesExistingFilterAndValues(string options)
    {
        var (_, name) = CreatePivot();
        using (Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "Region",
            filterOptions = new { type = "ValueIsGreaterThan", number1 = 150d, dataFieldName = "Total Sales" }
        })) { }
        using var request = JsonDocument.Parse(options);
        var failure = await _fixture.SendForFailureAsync("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "Region",
            filterOptions = request.RootElement
        });
        Assert.False(failure.Success);
        using var state = Send("pivottablefield.get-field-filters", new { pivotTableName = name, fieldName = "Region" });
        Assert.Equal("ValueIsGreaterThan", Assert.Single(state.RootElement.GetProperty("filters").EnumerateArray()).GetProperty("type").GetString());
        AssertTotal(name, 300d);
    }

    [Fact]
    public void MultipleFiltersAndClearing_PreserveOtherFieldsAndManualVisibility()
    {
        var (_, name) = CreatePivot();
        using (Send("pivottablecalc.set-layout-options", new { pivotTableName = name, layoutOptions = new { allowMultipleFilters = true } })) { }
        foreach (var type in new[] { "CaptionContains", "ValueIsGreaterThan" })
            using (Send("pivottablefield.add-field-filter", new
            {
                pivotTableName = name,
                fieldName = "Region",
                filterOptions = type == "CaptionContains" ? (object)new { type, text1 = "h" } :
                    new { type, number1 = 150d, dataFieldName = "Total Sales" }
            })) { }
        using (var state = Send("pivottablefield.get-field-filters", new { pivotTableName = name, fieldName = "Region" }))
            Assert.Equal(2, state.RootElement.GetProperty("filters").GetArrayLength());
        using (Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "Product",
            filterOptions = new { type = "CaptionEquals", text1 = "B" }
        })) { }
        using (Send("pivottablefield.clear-field-filters", new { pivotTableName = name, fieldName = "Region" })) { }
        using (var state = Send("pivottablefield.get-field-filters", new { pivotTableName = name, fieldName = "Product" }))
            Assert.Single(state.RootElement.GetProperty("filters").EnumerateArray());
        AssertTotal(name, 240d);
        using (Send("pivottablefield.clear-field-filters", new { pivotTableName = name, fieldName = "Product" })) { }
        using (Send("pivottablefield.set-field-filter", new { pivotTableName = name, fieldName = "Region", selectedValues = new List<string> { "North" } })) { }
        using (Send("pivottablefield.clear-field-filters", new { pivotTableName = name, fieldName = "Region" })) { }
        AssertTotal(name, 100d);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public void LayoutForms_ReadEveryNativeRowField(int rowLayout, bool repeatLabels)
    {
        var (_, name) = CreatePivot();
        using var state = Send("pivottablecalc.set-layout-options", new
        {
            pivotTableName = name,
            layoutOptions = new { rowLayout, repeatLabels, showColumnStripes = true, showRowHeaders = false, showColumnHeaders = false }
        });
        var rows = state.RootElement.GetProperty("rowFields").EnumerateArray().ToArray();
        Assert.Equal(2, rows.Length);
        Assert.All(rows, row =>
        {
            Assert.Equal(rowLayout, row.GetProperty("rowLayout").GetInt32());
            Assert.Equal(repeatLabels, row.GetProperty("repeatLabels").GetBoolean());
        });
        Assert.True(state.RootElement.GetProperty("showColumnStripes").GetBoolean());
        Assert.False(state.RootElement.GetProperty("showRowHeaders").GetBoolean());
        Assert.False(state.RootElement.GetProperty("showColumnHeaders").GetBoolean());
    }

    [Theory]
    [InlineData("""{"rowLayout":3,"showRowStripes":true}""")]
    [InlineData("""{"rowLayout":1,"styleName":"missing-style","showRowStripes":true}""")]
    [InlineData("""{"rowLayout":0,"repeatLabels":true,"showRowStripes":true}""")]
    [InlineData("""{"repeatLabels":true,"showRowStripes":true}""")]
    [InlineData("""{"unknown":true}""")]
    public async Task InvalidLayout_PreservesEveryOriginalSetting(string options)
    {
        var (_, name) = CreatePivot();
        using var before = Send("pivottablecalc.get-layout-options", new { pivotTableName = name });
        using var request = JsonDocument.Parse(options);
        var failure = await _fixture.SendForFailureAsync("pivottablecalc.set-layout-options", new
        {
            pivotTableName = name,
            layoutOptions = request.RootElement
        });
        Assert.False(failure.Success);
        using var after = Send("pivottablecalc.get-layout-options", new { pivotTableName = name });
        Assert.Equal(before.RootElement.GetRawText(), after.RootElement.GetRawText());
    }

    [Fact]
    public void SourceReplacement_IsolatesSelectedPivotAndPreservesOtherSharedSlicer()
    {
        var (sheet, name) = CreatePivot();
        var other = CreateSharedPivot(sheet, name);
        string slicer = $"Control_{Guid.NewGuid():N}";
        using (Send("slicer.create-slicer", new { pivotTableName = other, fieldName = "Region", slicerName = slicer, destinationSheet = sheet, position = "N2" })) { }
        using var original = Send("pivottable.get-source", new { pivotTableName = other });
        Assert.Equal(2, original.RootElement.GetProperty("sharedPivotTables").GetArrayLength());
        Assert.Single(original.RootElement.GetProperty("connectedSlicerCaches").EnumerateArray());
        using (Send("range.set-values", new { sheetName = sheet, rangeAddress = "A8:C10", values = new object[][] { ["Region", "Product", "Sales"], ["North", "A", 250], ["South", "B", 450] } })) { }
        using var replacement = Send("pivottable.set-source", new { pivotTableName = name, sourceSheetName = sheet, sourceRangeAddress = "A8:C10" });
        Assert.NotEqual(original.RootElement.GetProperty("cacheIndex").GetInt32(), replacement.RootElement.GetProperty("cacheIndex").GetInt32());
        Assert.Equal(name, Assert.Single(replacement.RootElement.GetProperty("sharedPivotTables").EnumerateArray()).GetString());
        AssertTotal(name, 700d);
        AssertTotal(other, 400d);
        using var unchanged = Send("pivottable.get-source", new { pivotTableName = other });
        Assert.Equal(original.RootElement.GetProperty("sourceData").GetString(), unchanged.RootElement.GetProperty("sourceData").GetString());
        Assert.Equal(original.RootElement.GetProperty("cacheIndex").GetInt32(), unchanged.RootElement.GetProperty("cacheIndex").GetInt32());
        using var control = Send("slicer.get-slicer", new { slicerName = slicer });
        Assert.Equal(other, Assert.Single(control.RootElement.GetProperty("slicer").GetProperty("connectedPivotTables").EnumerateArray()).GetString());
    }

    [Fact]
    public async Task ConnectedSlicer_RejectsSourceReplacementAndPreservesConnection()
    {
        var (sheet, name) = CreatePivot();
        string slicer = $"Control_{Guid.NewGuid():N}";
        using (Send("slicer.create-slicer", new { pivotTableName = name, fieldName = "Region", slicerName = slicer, destinationSheet = sheet, position = "N2" })) { }
        using var original = Send("pivottable.get-source", new { pivotTableName = name });
        var failure = await _fixture.SendForFailureAsync("pivottable.set-source", new { pivotTableName = name, sourceSheetName = sheet, sourceRangeAddress = "A1:C5" });
        Assert.False(failure.Success);
        Assert.Contains("Disconnect", failure.ErrorMessage, StringComparison.Ordinal);
        using var after = Send("pivottable.get-source", new { pivotTableName = name });
        Assert.Equal(original.RootElement.GetRawText(), after.RootElement.GetRawText());
    }

    [Fact]
    public async Task SaveReopen_PreservesSourceStyleRepeatedLabelsFilterAndExpansion()
    {
        var (sheet, name) = CreatePivot();
        using (Send("range.set-values", new { sheetName = sheet, rangeAddress = "A8:C12", values = new object[][] { ["Region", "Product", "Sales"], ["North", "A", 40], ["North", "B", 60], ["South", "A", 120], ["South", "B", 180] } })) { }
        using (Send("pivottable.set-source", new { pivotTableName = name, sourceSheetName = sheet, sourceRangeAddress = "A8:C12" })) { }
        using (Send("pivottablecalc.set-layout-options", new { pivotTableName = name, layoutOptions = new { rowLayout = 1, repeatLabels = true, styleName = "PivotStyleMedium9", preserveFormatting = true } })) { }
        using (Send("pivottablefield.add-field-filter", new { pivotTableName = name, fieldName = "Region", filterOptions = new { type = "CaptionEquals", text1 = "North" } })) { }
        using (Send("pivottablefield.set-item-expansion", new { pivotTableName = name, fieldName = "Region", itemName = "North", expanded = false })) { }
        await _fixture.SaveAndReopenAsync();
        using var source = Send("pivottable.get-source", new { pivotTableName = name });
        Assert.Contains("R8C1", source.RootElement.GetProperty("sourceData").GetString(), StringComparison.Ordinal);
        using var layout = Send("pivottablecalc.get-layout-options", new { pivotTableName = name });
        Assert.Equal("PivotStyleMedium9", layout.RootElement.GetProperty("styleName").GetString());
        Assert.All(layout.RootElement.GetProperty("rowFields").EnumerateArray(), field => Assert.True(field.GetProperty("repeatLabels").GetBoolean()));
        using var filters = Send("pivottablefield.get-field-filters", new { pivotTableName = name, fieldName = "Region" });
        Assert.Equal("North", Assert.Single(filters.RootElement.GetProperty("filters").EnumerateArray()).GetProperty("value1").GetString());
        using var item = Send("pivottablefield.get-item-expansion", new { pivotTableName = name, fieldName = "Region", itemName = "North" });
        Assert.False(item.RootElement.GetProperty("expanded").GetBoolean());
        AssertTotal(name, 100d);
    }

    private void AssertTotal(string name, double expected)
    {
        var data = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(data.Success, data.ErrorMessage);
        Assert.Equal(expected, Convert.ToDouble(data.Values[^1][^1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("A8:C9")]
    [InlineData("A8:C8")]
    [InlineData("A8:B9")]
    [InlineData("A8:C9,A11:C12")]
    [InlineData("E1:G5")]
    public async Task InvalidSource_PreservesOriginalCacheAndLayout(string address)
    {
        var (sheet, name) = CreatePivot();
        using (Send("range.set-values", new { sheetName = sheet, rangeAddress = "A8:C9", values = new object[][] { ["Region", "Product", "Wrong"], ["North", "A", 250] } })) { }
        using var before = Send("pivottable.get-source", new { pivotTableName = name });
        using var layoutBefore = Send("pivottablecalc.get-layout-options", new { pivotTableName = name });
        var failure = await _fixture.SendForFailureAsync("pivottable.set-source", new { pivotTableName = name, sourceSheetName = sheet, sourceRangeAddress = address });
        Assert.False(failure.Success);
        using var after = Send("pivottable.get-source", new { pivotTableName = name });
        using var layoutAfter = Send("pivottablecalc.get-layout-options", new { pivotTableName = name });
        Assert.Equal(before.RootElement.GetRawText(), after.RootElement.GetRawText());
        Assert.Equal(layoutBefore.RootElement.GetRawText(), layoutAfter.RootElement.GetRawText());
        AssertTotal(name, 400d);
    }

    [Fact]
    public void TableSource_ExcludesTotalsAndKeepsNativeTableReference()
    {
        var (sheet, name) = CreatePivot();
        var table = $"Source_{Guid.NewGuid():N}";
        using (Send("range.set-values", new { sheetName = sheet, rangeAddress = "A8:C10", values = new object[][] { ["Region", "Product", "Sales"], ["North", "A", 250], ["South", "B", 450] } })) { }
        using (Send("table.create", new { sheetName = sheet, tableName = table, rangeAddress = "A8:C10" })) { }
        using (Send("table.toggle-totals", new { tableName = table, showTotals = true })) { }
        using var replaced = Send("pivottable.set-source", new { pivotTableName = name, sourceSheetName = sheet, tableName = table });
        Assert.Equal(table, replaced.RootElement.GetProperty("sourceData").GetString());
        Assert.Equal(2, replaced.RootElement.GetProperty("recordCount").GetInt32());
        AssertTotal(name, 700d);
    }

    [Fact]
    public async Task SourceReplacement_PreservesRefreshAndRetentionSettings()
    {
        var (sheet, name) = CreatePivot();
        using (Send("pivottable.set-cache-options", new { pivotTableName = name, refreshOnFileOpen = true, missingItemsLimit = "None", saveSourceData = false })) { }
        using (Send("pivottable.set-source", new { pivotTableName = name, sourceSheetName = sheet, sourceRangeAddress = "A1:C5" })) { }
        using var after = Send("pivottable.get-cache-options", new { pivotTableName = name });
        Assert.True(after.RootElement.GetProperty("refreshOnFileOpen").GetBoolean());
        Assert.Equal("None", after.RootElement.GetProperty("missingItemsLimit").GetString());
        Assert.False(after.RootElement.GetProperty("saveSourceData").GetBoolean());
        using (Send("pivottable.set-cache-options", new { pivotTableName = name, enableRefresh = false })) { }
        var failure = await _fixture.SendForFailureAsync("pivottable.set-source", new { pivotTableName = name, sourceSheetName = sheet, sourceRangeAddress = "A1:C5" });
        Assert.False(failure.Success);
        Assert.Contains("Enable cache refresh", failure.ErrorMessage, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Product", "A")]
    [InlineData("Region", "Missing")]
    [InlineData("Sales", "100")]
    public async Task InvalidExpansion_PreservesParentAndOtherItem(string fieldName, string itemName)
    {
        var (_, name) = CreatePivot();
        using var before = Send("pivottablefield.get-item-expansion", new { pivotTableName = name, fieldName = "Region", itemName = "North" });
        var failure = await _fixture.SendForFailureAsync("pivottablefield.set-item-expansion", new { pivotTableName = name, fieldName, itemName, expanded = false });
        Assert.False(failure.Success);
        using var after = Send("pivottablefield.get-item-expansion", new { pivotTableName = name, fieldName = "Region", itemName = "North" });
        Assert.Equal(before.RootElement.GetRawText(), after.RootElement.GetRawText());
        using var other = Send("pivottablefield.get-item-expansion", new { pivotTableName = name, fieldName = "Region", itemName = "South" });
        Assert.True(other.RootElement.GetProperty("expanded").GetBoolean());
    }

    private string CreateSharedPivot(string sheetName, string name)
    {
        var otherName = $"Shared_{Guid.NewGuid():N}";
        // Capability setup: the public create actions intentionally allocate separate caches.
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.PivotTable? original = null;
            Excel.PivotCache? cache = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? anchor = null;
            Excel.PivotTable? created = null;
            try
            {
                original = CoreLookupHelpers.FindPivotTable(ctx.Book, name);
                cache = original.PivotCache();
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                anchor = sheet.Range["J1"];
                created = cache.CreatePivotTable(anchor, otherName);
            }
            finally
            {
                ComUtilities.Release(ref created);
                ComUtilities.Release(ref anchor);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref original);
            }
        });
        var row = _pivot.AddRowField(_fixture.BatchToken, otherName, "Region");
        Assert.True(row.Success, row.ErrorMessage);
        var value = _pivot.AddValueField(_fixture.BatchToken, otherName, "Sales", AggregationFunction.Sum, "Total Sales");
        Assert.True(value.Success, value.ErrorMessage);
        return otherName;
    }

    private (string Sheet, string Name) CreatePivot()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Depth_{Guid.NewGuid():N}";
        var written = _commands.SetValues(_fixture.BatchToken, sheet, "A1:C5",
            [["Region", "Product", "Sales"], ["North", "A", 40], ["North", "B", 60],
                ["South", "A", 120], ["South", "B", 180]]);
        Assert.True(written.Success, written.ErrorMessage);
        var created = _pivot.CreateFromRange(_fixture.BatchToken, sheet, "A1:C5", sheet, "E1", name);
        Assert.True(created.Success, created.ErrorMessage);
        var row = _pivot.AddRowField(_fixture.BatchToken, name, "Region");
        Assert.True(row.Success, row.ErrorMessage);
        var child = _pivot.AddRowField(_fixture.BatchToken, name, "Product");
        Assert.True(child.Success, child.ErrorMessage);
        var value = _pivot.AddValueField(_fixture.BatchToken, name, "Sales", AggregationFunction.Sum, "Total Sales");
        Assert.True(value.Success, value.ErrorMessage);
        return (sheet, name);
    }

    private void AssertValues(string name, double total, params (string Label, double Amount)[] expected)
    {
        var data = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(data.Success, data.ErrorMessage);
        Assert.Equal(expected.Length + 2, data.Values.Count);
        Assert.All(data.Values, row => Assert.Equal(2, row.Count));
        Assert.Equal(expected, data.Values.Skip(1).SkipLast(1).Select(row =>
            (Label: Assert.IsType<string>(row[0]),
                Amount: Convert.ToDouble(row[1], System.Globalization.CultureInfo.InvariantCulture))));
        Assert.Equal(total, Convert.ToDouble(data.Values[^1][^1], System.Globalization.CultureInfo.InvariantCulture));
    }

    private JsonDocument Send(string action, object arguments)
    {
        var response = _fixture.Send(action, arguments);
        Assert.NotNull(response.Result);
        var document = JsonDocument.Parse(response.Result);
        Assert.True(document.RootElement.GetProperty("success").GetBoolean());
        return document;
    }
}
