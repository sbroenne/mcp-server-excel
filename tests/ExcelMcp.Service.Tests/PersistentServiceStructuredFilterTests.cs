using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "StructuredFilters")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceStructuredFilterTests(PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceWorkbookFixture>
{
    [Fact]
    public void OrdinaryRange_CombinesBothConditionsAndReturnsEveryColumn()
    {
        var sheet = CreateData();
        _fixture.Send("rangeedit.apply-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            columnIndex = 2,
            filterOptions = new { filterOperator = "And", criteria1 = ">=20", criteria2 = "<=40" }
        });
        var response = _fixture.Send("rangeedit.get-filters", new { sheetName = sheet, rangeAddress = "A1:B6" });
        using var read = JsonDocument.Parse(response.Result!);
        var filters = read.RootElement.GetProperty("columnFilters");
        Assert.Equal(2, filters.GetArrayLength());
        Assert.False(filters[0].GetProperty("isFiltered").GetBoolean());
        Assert.Equal("And", filters[1].GetProperty("filterOperator").GetString());
        Assert.Equal(">=20", filters[1].GetProperty("criteria1").GetProperty("value").GetString());
        Assert.Equal("<=40", filters[1].GetProperty("criteria2").GetProperty("value").GetString());
        AssertVisibleRows(sheet, 3, 4, 5);
    }

    [Fact]
    public void Table_ReadPreservesNativeValueArrays()
    {
        var sheet = CreateData();
        var table = "Filter_" + Guid.NewGuid().ToString("N");
        _fixture.Send("table.create", new { sheetName = sheet, tableName = table, rangeAddress = "A1:B6", hasHeaders = true });
        string[] values = ["A", "C", "not-present"];
        _fixture.Send("tablecolumn.apply-filter", new
        {
            tableName = table,
            columnName = "Category",
            options = new { filterOperator = "Values", values }
        });
        var response = _fixture.Send("tablecolumn.get-filters", new { tableName = table });
        using var read = JsonDocument.Parse(response.Result!);
        var first = read.RootElement.GetProperty("columnFilters")[0];
        Assert.Equal("Values", first.GetProperty("filterOperator").GetString());
        var criteria = first.GetProperty("criteria1");
        Assert.True(criteria.GetProperty("available").GetBoolean());
        Assert.Equal(values.Select(value => "=" + value),
            criteria.GetProperty("value").EnumerateArray().Select(item => item.GetString()));
        AssertVisibleRows(sheet, 2, 4, 5);
    }

    [Theory]
    [InlineData(false, "Comparison", """{"criteria1":">=30"}""", new[] { 4, 5, 6 })]
    [InlineData(true, "Comparison", """{"criteria1":">=30"}""", new[] { 4, 5, 6 })]
    [InlineData(false, "Or", """{"criteria1":"<=10","criteria2":">=50"}""", new[] { 2, 6 })]
    [InlineData(true, "Or", """{"criteria1":"<=10","criteria2":">=50"}""", new[] { 2, 6 })]
    [InlineData(false, "TopItems", """{"count":2}""", new[] { 5, 6 })]
    [InlineData(true, "TopItems", """{"count":2}""", new[] { 5, 6 })]
    [InlineData(false, "BottomItems", """{"count":2}""", new[] { 2, 3 })]
    [InlineData(true, "BottomItems", """{"count":2}""", new[] { 2, 3 })]
    [InlineData(false, "TopPercent", """{"count":40}""", new[] { 5, 6 })]
    [InlineData(true, "TopPercent", """{"count":40}""", new[] { 5, 6 })]
    [InlineData(false, "BottomPercent", """{"count":40}""", new[] { 2, 3 })]
    [InlineData(true, "BottomPercent", """{"count":40}""", new[] { 2, 3 })]
    [InlineData(false, "Dynamic", """{"dynamicCriteria":"xlFilterAboveAverage"}""", new[] { 5, 6 })]
    [InlineData(true, "Dynamic", """{"dynamicCriteria":"xlFilterAboveAverage"}""", new[] { 5, 6 })]
    [InlineData(false, "Dynamic", """{"dynamicCriteria":"xlFilterBelowAverage"}""", new[] { 2, 3 })]
    [InlineData(true, "Dynamic", """{"dynamicCriteria":"xlFilterBelowAverage"}""", new[] { 2, 3 })]
    public void NativeNumericFilters_ApplyAndClearOnlyTheirOwnScope(bool tableMode, string kind, string json, int[] expectedRows)
    {
        var sheet = CreateData();
        string? table = tableMode ? CreateTable(sheet) : null;
        var otherSheet = CreateData();
        var otherTable = CreateTable(otherSheet);
        Apply(otherSheet, otherTable, 2, new { criteria1 = "=20" });
        var original = _commands.GetValues(_fixture.BatchToken, sheet, "A1:B6");
        Assert.True(original.Success, original.ErrorMessage);
        using var extra = JsonDocument.Parse(json);
        var options = extra.RootElement.EnumerateObject().ToDictionary(property => property.Name,
            property => (object?)property.Value.Clone(), StringComparer.Ordinal);
        options["filterOperator"] = kind;
        Apply(sheet, table, 2, options);
        AssertVisibleRows(sheet, expectedRows);
        AssertVisibleRows(otherSheet, 3);
        _fixture.Send(tableMode ? "tablecolumn.clear-filters" : "rangeedit.clear-filters",
            tableMode ? new { tableName = table } : (object)new { sheetName = sheet, rangeAddress = "A1:B6" });
        AssertVisibleRows(sheet, 2, 3, 4, 5, 6);
        AssertVisibleRows(otherSheet, 3);
        _fixture.Send(tableMode ? "tablecolumn.clear-filters" : "rangeedit.clear-filters",
            tableMode ? new { tableName = table } : (object)new { sheetName = sheet, rangeAddress = "A1:B6" });
        AssertVisibleRows(sheet, 2, 3, 4, 5, 6);
        AssertVisibleRows(otherSheet, 3);
        var after = _commands.GetValues(_fixture.BatchToken, sheet, "A1:B6");
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(original.Values), JsonSerializer.Serialize(after.Values));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeFilters_ComposeAcrossColumns(bool tableMode)
    {
        var sheet = CreateData();
        string? table = tableMode ? CreateTable(sheet) : null;
        string[] values = ["A", "C", "not-present"];
        Apply(sheet, table, 1, new { filterOperator = "Values", values });
        Apply(sheet, table, 2, new { criteria1 = ">=30" });
        AssertVisibleRows(sheet, 4, 5);
    }

    [Theory]
    [InlineData("Year", "2026-02-01", new[] { 3, 4, 5 })]
    [InlineData("Month", "2026-02-01", new[] { 4, 5 })]
    [InlineData("Day", "2026-02-03", new[] { 4 })]
    public void DateGroups_UseNativeCalendarGroups(string level, string date, int[] expectedRows)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B6",
            [["Category", "Date"], ["A", new DateTime(2025, 12, 1).ToOADate()],
                ["B", new DateTime(2026, 1, 2).ToOADate()], ["C", new DateTime(2026, 2, 3).ToOADate()],
                ["A", new DateTime(2026, 2, 4).ToOADate()], ["B", new DateTime(2027, 1, 1).ToOADate()]]).Success);
        Assert.True(_commands.SetNumberFormat(_fixture.BatchToken, sheet, "B2:B6", "yyyy-mm-dd").Success);
        Apply(sheet, null, 2, new
        {
            filterOperator = "Values",
            dateGroups = new[] { new { level, date } }
        });
        AssertVisibleRows(sheet, expectedRows);
        var response = _fixture.Send("rangeedit.get-filters", new { sheetName = sheet, rangeAddress = "A1:B6" });
        using var read = JsonDocument.Parse(response.Result!);
        var column = read.RootElement.GetProperty("columnFilters")[1];
        Assert.True(column.GetProperty("isFiltered").GetBoolean());
        Assert.True(column.TryGetProperty("criteria1", out _));
        Assert.True(column.TryGetProperty("criteria2", out _));
    }

    [Fact]
    public async Task OrdinaryFilter_DifferentScopeDoesNotReplaceAnExistingFilter()
    {
        var sheet = CreateData();
        Apply(sheet, null, 2, new { criteria1 = ">=30" });
        var response = await _fixture.SendForFailureAsync("rangeedit.apply-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B5",
            columnIndex = 2,
            filterOptions = new { criteria1 = "<=20" }
        });
        Assert.False(response.Success);
        Assert.Contains("different range", response.ErrorMessage);
        AssertVisibleRows(sheet, 4, 5, 6);
    }

    [Fact]
    public async Task OrdinaryFilter_RequiresExplicitClearingOfAnExistingAdvancedFilter()
    {
        var sheet = CreateData();
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "H1:H2", [["Amount"], [">=30"]]).Success);
        var original = _commands.GetValues(_fixture.BatchToken, sheet, "A1:B6");
        Assert.True(original.Success, original.ErrorMessage);
        _fixture.Send("rangeedit.advanced-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            criteriaRange = "H1:H2",
            mode = "InPlace"
        });
        AssertVisibleRows(sheet, 4, 5, 6);
        var before = _fixture.Send("rangeedit.get-filters", new { sheetName = sheet, rangeAddress = "A1:B6" });
        using (var state = JsonDocument.Parse(before.Result!))
        {
            Assert.True(state.RootElement.GetProperty("worksheetFilterMode").GetBoolean());
            Assert.False(state.RootElement.GetProperty("filterEnabled").GetBoolean());
            Assert.False(state.RootElement.GetProperty("advancedCriteriaAvailable").GetBoolean());
        }
        var rejected = await _fixture.SendForFailureAsync("rangeedit.apply-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            columnIndex = 2,
            filterOptions = new { criteria1 = "<=20" }
        });
        Assert.False(rejected.Success);
        Assert.Contains("clear_advanced", rejected.ErrorMessage);
        Assert.Contains("--clear-advanced", rejected.ErrorMessage);
        AssertVisibleRows(sheet, 4, 5, 6);
        var after = _fixture.Send("rangeedit.get-filters", new { sheetName = sheet, rangeAddress = "A1:B6" });
        Assert.Equal(before.Result, after.Result);
        var unchanged = _commands.GetValues(_fixture.BatchToken, sheet, "A1:B6");
        Assert.True(unchanged.Success, unchanged.ErrorMessage);
        Assert.Equal(JsonSerializer.Serialize(original.Values), JsonSerializer.Serialize(unchanged.Values));
        var criteria = _commands.GetValues(_fixture.BatchToken, sheet, "H1:H2");
        Assert.True(criteria.Success, criteria.ErrorMessage);
        Assert.Equal("Amount", criteria.Values[0][0]);
        Assert.Equal(">=30", criteria.Values[1][0]);

        _fixture.Send("rangeedit.clear-filters",
            new { sheetName = sheet, rangeAddress = "A1:B6", clearAdvanced = true });
        AssertVisibleRows(sheet, 2, 3, 4, 5, 6);
        Apply(sheet, null, 2, new { criteria1 = "<=20" });
        AssertVisibleRows(sheet, 2, 3);
        var applied = _fixture.Send("rangeedit.get-filters", new { sheetName = sheet, rangeAddress = "A1:B6" });
        using var ordinary = JsonDocument.Parse(applied.Result!);
        Assert.True(ordinary.RootElement.GetProperty("filterEnabled").GetBoolean());
        Assert.True(ordinary.RootElement.GetProperty("advancedCriteriaAvailable").GetBoolean());
        Assert.Equal("<=20", ordinary.RootElement.GetProperty("columnFilters")[1]
            .GetProperty("criteria1").GetProperty("value").GetString());
    }

    [Fact]
    public void AdvancedFilter_CopiesMatchingRecordsWithoutModifyingSource()
    {
        var sheet = CreateData();
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "H1:H2", [["Amount"], [">=30"]]).Success);
        _fixture.Send("rangeedit.advanced-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            criteriaRange = "H1:H2",
            mode = "Copy",
            copyToRange = "D1",
            uniqueOnly = false
        });
        var result = _commands.GetValues(_fixture.BatchToken, sheet, "D1:E5");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("Category", result.Values[0][0]);
        Assert.Equal(30d, Convert.ToDouble(result.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal("C", result.Values[1][0]);
        Assert.Equal("A", result.Values[2][0]);
        Assert.Equal(40d, Convert.ToDouble(result.Values[2][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal("B", result.Values[3][0]);
        Assert.Equal(50d, Convert.ToDouble(result.Values[3][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Null(result.Values[4][0]);
        AssertVisibleRows(sheet, 2, 3, 4, 5, 6);
        var source = _commands.GetValues(_fixture.BatchToken, sheet, "A1:B6");
        Assert.True(source.Success, source.ErrorMessage);
        Assert.Equal([10d, 20d, 30d, 40d, 50d],
            source.Values.Skip(1).Select(row => Convert.ToDouble(row[1], System.Globalization.CultureInfo.InvariantCulture)));
    }

    [Theory]
    [InlineData("CellColor", "fillColor")]
    [InlineData("FontColor", "fontColor")]
    public void ColorFilters_UseNativeDisplayedColors(string kind, string property)
    {
        var sheet = CreateData();
        var format = new Dictionary<string, object?>
        {
            ["sheetName"] = sheet,
            ["rangeAddresses"] = (string[])["B2:B3"],
            ["formatOptions"] = new Dictionary<string, object?> { [property] = "#FF0000" }
        };
        _fixture.Send("rangeformat.format", format);
        Apply(sheet, null, 2, new { filterOperator = kind, color = "#FF0000" });
        AssertVisibleRows(sheet, 2, 3);
        if (kind == "CellColor")
        {
            _fixture.ExecuteRawVerification((context, _) =>
            {
                Excel.Worksheet? worksheet = null;
                Excel.AutoFilter? filter = null;
                Excel.Filters? filters = null;
                Excel.Filter? column = null;
                object? criterion = null;
                try
                {
                    worksheet = ComUtilities.FindSheet(context.Book, sheet);
                    filter = worksheet!.AutoFilter;
                    filters = filter.Filters;
                    column = filters[2];
                    criterion = column.Criteria1;
                    Assert.True(criterion is Excel.Interior,
                        $"Interior={criterion is Excel.Interior}; FormatColor={criterion is Excel.FormatColor}; Font={criterion is Excel.Font}; Range={criterion is Excel.Range}");
                }
                finally
                {
                    ComUtilities.Release(ref criterion);
                    ComUtilities.Release(ref column);
                    ComUtilities.Release(ref filters);
                    ComUtilities.Release(ref filter);
                    ComUtilities.Release(ref worksheet);
                }
            });
        }
        var response = _fixture.Send("rangeedit.get-filters", new { sheetName = sheet, rangeAddress = "A1:B6" });
        using var read = JsonDocument.Parse(response.Result!);
        var column = read.RootElement.GetProperty("columnFilters")[1];
        Assert.Equal(kind, column.GetProperty("filterOperator").GetString());
        Assert.Equal(255, column.GetProperty("criteria1").GetProperty("value").GetInt32());
    }

    [Fact]
    public void IconFilter_ReadsTheActualNativeIconIdentity()
    {
        var sheet = CreateData();
        _fixture.Send("conditionalformat.add-rule", new
        {
            sheetName = sheet,
            rangeAddress = "B2:B6",
            ruleType = "iconSet",
            iconSetId = "3TrafficLights1"
        });
        Apply(sheet, null, 2, new { filterOperator = "Icon", iconSet = "xl3TrafficLights1", iconIndex = 3 });
        AssertVisibleRows(sheet, 5, 6);
        var response = _fixture.Send("rangeedit.get-filters", new { sheetName = sheet, rangeAddress = "A1:B6" });
        using var read = JsonDocument.Parse(response.Result!);
        var icon = read.RootElement.GetProperty("columnFilters")[1].GetProperty("criteria1").GetProperty("value");
        Assert.Equal("xl3TrafficLights1", icon.GetProperty("iconSet").GetString());
        Assert.Equal(3, icon.GetProperty("iconIndex").GetInt32());
    }

    [Fact]
    public void AdvancedFilter_InPlaceCanBeClearedWithoutLosingRecords()
    {
        var sheet = CreateData();
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "H1:H2", [["Amount"], [">=30"]]).Success);
        _fixture.Send("rangeedit.advanced-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            criteriaRange = "H1:H2",
            mode = "InPlace"
        });
        AssertVisibleRows(sheet, 4, 5, 6);
        _fixture.Send("rangeedit.clear-filters", new { sheetName = sheet, rangeAddress = "A1:B6", clearAdvanced = true });
        AssertVisibleRows(sheet, 2, 3, 4, 5, 6);
    }

    [Fact]
    public async Task AdvancedFilter_RejectsOccupiedMaximumCopyExtentBeforeWriting()
    {
        var sheet = CreateData();
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "H1:H2", [["Amount"], [">=30"]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "E6", [["keep"]]).Success);
        var response = await _fixture.SendForFailureAsync("rangeedit.advanced-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            criteriaRange = "H1:H2",
            mode = "Copy",
            copyToRange = "D1"
        });
        Assert.False(response.Success);
        Assert.Contains("$E$6", response.ErrorMessage);
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "D1:E6");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Null(read.Values[0][0]);
        Assert.Equal("keep", read.Values[5][1]);
    }

    [Theory]
    [InlineData("""{"filterOperator":"And","criteria1":">=20"}""")]
    [InlineData("""{"filterOperator":"TopItems","count":0}""")]
    [InlineData("""{"filterOperator":"TopPercent","count":101}""")]
    [InlineData("""{"filterOperator":"Values","values":[]}""")]
    [InlineData("""{"filterOperator":"Values","values":["A"],"criteria1":"=A"}""")]
    [InlineData("""{"filterOperator":"CellColor","color":"invalid"}""")]
    [InlineData("""{"filterOperator":"Icon","iconSet":"wrong","iconIndex":1}""")]
    [InlineData("""{"filterOperator":"Dynamic","dynamicCriteria":"wrong"}""")]
    [InlineData("""{"criteria1":">=20","unknown":true}""")]
    public async Task InvalidSettings_DoNotChangeAnExistingFilter(string json)
    {
        var sheet = CreateData();
        Apply(sheet, null, 2, new { criteria1 = ">=30" });
        using var options = JsonDocument.Parse(json);
        var response = await _fixture.SendForFailureAsync("rangeedit.apply-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            columnIndex = 2,
            filterOptions = options.RootElement
        });
        Assert.False(response.Success);
        AssertVisibleRows(sheet, 4, 5, 6);
    }

    [Fact]
    public async Task AdvancedFilter_ClearRequiresExplicitWorksheetWidePermission()
    {
        var sheet = CreateData();
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "H1:H2", [["Amount"], [">=30"]]).Success);
        _fixture.Send("rangeedit.advanced-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            criteriaRange = "H1:H2",
            mode = "InPlace"
        });
        var read = _fixture.Send("rangeedit.get-filters", new { sheetName = sheet, rangeAddress = "A1:B6" });
        using var state = JsonDocument.Parse(read.Result!);
        Assert.True(state.RootElement.GetProperty("worksheetFilterMode").GetBoolean());
        Assert.False(state.RootElement.GetProperty("advancedCriteriaAvailable").GetBoolean());
        var response = await _fixture.SendForFailureAsync("rangeedit.clear-filters",
            new { sheetName = sheet, rangeAddress = "A1:B6" });
        Assert.False(response.Success);
        Assert.Contains("clear_advanced", response.ErrorMessage);
        AssertVisibleRows(sheet, 4, 5, 6);
    }

    [Fact]
    public void AdvancedFilter_UniqueCopyKeepsOnlyDistinctRequestedFields()
    {
        var sheet = CreateData();
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "H1:H2", [["Amount"], [">=10"]]).Success);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "D1", [["Category"]]).Success);
        _fixture.Send("rangeedit.advanced-filter", new
        {
            sheetName = sheet,
            rangeAddress = "A1:B6",
            criteriaRange = "H1:H2",
            mode = "Copy",
            copyToRange = "D1",
            uniqueOnly = true,
            overwritePolicy = "allow"
        });
        var read = _commands.GetValues(_fixture.BatchToken, sheet, "D1:D5");
        Assert.True(read.Success, read.ErrorMessage);
        Assert.Equal("Category", read.Values[0][0]);
        Assert.Equal(["A", "B", "C"], read.Values.Skip(1).Take(3).Select(row => row[0]));
        Assert.Null(read.Values[4][0]);
    }

    private string CreateTable(string sheet)
    {
        var table = "Filter_" + Guid.NewGuid().ToString("N");
        _fixture.Send("table.create", new { sheetName = sheet, tableName = table, rangeAddress = "A1:B6", hasHeaders = true });
        return table;
    }

    private void Apply(string sheet, string? table, int column, object options)
    {
        if (table is not null)
            _fixture.Send("tablecolumn.apply-filter", new { tableName = table, columnName = column == 1 ? "Category" : "Amount", options });
        else
            _fixture.Send("rangeedit.apply-filter", new
            {
                sheetName = sheet,
                rangeAddress = "A1:B6",
                columnIndex = column,
                filterOptions = options
            });
    }

    private void AssertVisibleRows(string sheet, params int[] expectedRows)
    {
        var actualRows = _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? worksheet = null;
            Excel.Range? rows = null;
            Excel.Range? row = null;
            try
            {
                worksheet = ComUtilities.FindSheet(context.Book, sheet);
                rows = worksheet!.Rows;
                var visibleRows = new List<int>();
                for (int number = 2; number <= 6; number++)
                {
                    row = (Excel.Range)rows[number];
                    if (!Convert.ToBoolean(row.Hidden, System.Globalization.CultureInfo.InvariantCulture))
                        visibleRows.Add(number);
                    ComUtilities.Release(ref row);
                }
                return visibleRows.ToArray();
            }
            finally
            {
                ComUtilities.Release(ref row);
                ComUtilities.Release(ref rows);
                ComUtilities.Release(ref worksheet);
            }
        });
        Assert.Equal(expectedRows, actualRows);
        var visible = _fixture.Send("range.get-special-cells", new
        {
            sheetName = sheet,
            rangeAddress = "A2:A6",
            cellKind = "Visible"
        });
        using var read = JsonDocument.Parse(visible.Result!);
        Assert.Equal(expectedRows.Length, read.RootElement.GetProperty("cellCount").GetInt32());
    }

    private string CreateData()
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B6",
            [["Category", "Amount"], ["A", 10], ["B", 20], ["C", 30], ["A", 40], ["B", 50]]).Success);
        return sheet;
    }
}
