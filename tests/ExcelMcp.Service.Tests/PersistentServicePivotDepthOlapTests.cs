using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotDepth")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePivotDepthOlapTests(PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceDataModelFixture>
{
    [Theory]
    [InlineData("pivottablefield.get-item-expansion")]
    [InlineData("pivottablefield.set-item-expansion")]
    [InlineData("pivottable.get-source")]
    [InlineData("pivottable.set-source")]
    public async Task UnsupportedOlapMutation_FailsBeforeChangingData(string action)
    {
        var before = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = "DataModelPivot" });
        AssertSuccessful(before);
        AssertRegionalData(before.Result!);
        var args = new Dictionary<string, object?> { ["pivotTableName"] = "DataModelPivot" };
        if (action.StartsWith("pivottablefield.", StringComparison.Ordinal))
            args["fieldName"] = "[RegionalSalesTable].[Region]";
        if (action.Contains("item-expansion", StringComparison.Ordinal))
            args["itemName"] = "North";
        if (action.EndsWith("set-item-expansion", StringComparison.Ordinal))
            args["expanded"] = false;
        if (action.EndsWith("set-source", StringComparison.Ordinal))
        {
            args["sourceSheetName"] = "Sheet1";
            args["sourceRangeAddress"] = "A1:C5";
        }
        var failure = await _fixture.SendForFailureAsync(action, args);
        Assert.False(failure.Success);
        Assert.Equal(nameof(NotSupportedException), failure.ExceptionType);
        Assert.Contains(action.StartsWith("pivottablefield.", StringComparison.Ordinal)
                ? "This field operation supports regular PivotTables only"
                : "Source inspection/replacement supports worksheet-backed regular PivotTables only",
            failure.ErrorMessage, StringComparison.Ordinal);
        var after = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = "DataModelPivot" });
        AssertSuccessful(after);
        Assert.Equal(before.Result, after.Result);
        AssertRegionalData(after.Result!);
    }

    [Fact]
    public void OlapValueFiltersAndSort_ResolveMeasureCaptionAndUniqueName()
    {
        var pivot = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Filter_{Guid.NewGuid():N}";
        RequireSuccess(pivot.CreateFromDataModel(_fixture.BatchToken, "RegionalSalesTable", sheet, "A1", name));
        RequireSuccess(pivot.AddRowField(_fixture.BatchToken, name, "[RegionalSalesTable].[Region]", null));
        RequireSuccess(pivot.AddValueField(_fixture.BatchToken, name, "[Measures].[TotalRevenue]",
            Sbroenne.ExcelMcp.Core.Models.AggregationFunction.Sum, "Revenue"));
        var before = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(before);
        AssertRegionalData(before.Result!);

        var filtered = _fixture.Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "[RegionalSalesTable].[Region]",
            filterOptions = new { type = "TopCount", dataFieldName = "Revenue", number1 = 2 }
        });
        AssertSuccessful(filtered);
        using (var addedFilter = JsonDocument.Parse(filtered.Result!))
        {
            var filter = Assert.Single(addedFilter.RootElement.GetProperty("filters").EnumerateArray());
            Assert.Equal("TopCount", filter.GetProperty("type").GetString());
            Assert.Equal(2, filter.GetProperty("value1").GetInt32());
            Assert.Equal("Revenue", filter.GetProperty("dataFieldName").GetString());
        }

        var readFilters = _fixture.Send("pivottablefield.get-field-filters", new
        {
            pivotTableName = name,
            fieldName = "[RegionalSalesTable].[Region]"
        });
        AssertSuccessful(readFilters);
        using (var readFilter = JsonDocument.Parse(readFilters.Result!))
        {
            var filter = Assert.Single(readFilter.RootElement.GetProperty("filters").EnumerateArray());
            Assert.Equal("TopCount", filter.GetProperty("type").GetString());
            Assert.Equal(2, filter.GetProperty("value1").GetInt32());
            Assert.True(!filter.TryGetProperty("dataFieldName", out var dataFieldName) ||
                dataFieldName.ValueKind == JsonValueKind.Null);
        }
        var snapshot = RequireSuccess(pivot.Read(_fixture.BatchToken, name, sheet));
        Assert.False(snapshot.SnapshotComplete);
        var regionSnapshot = Assert.Single(snapshot.FieldSelections);
        Assert.False(regionSnapshot.FiltersAvailable);
        Assert.Contains("value field", regionSnapshot.FiltersError, StringComparison.OrdinalIgnoreCase);

        var topRows = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(topRows);
        AssertRegionalData(topRows.Result!, ("South", 12500d), ("West", 14500d));

        var cleared = _fixture.Send("pivottablefield.clear-field-filters", new
        {
            pivotTableName = name,
            fieldName = "[RegionalSalesTable].[Region]"
        });
        AssertSuccessful(cleared);
        using (var filters = JsonDocument.Parse(cleared.Result!))
            Assert.Empty(filters.RootElement.GetProperty("filters").EnumerateArray());

        var valueFiltered = _fixture.Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "[RegionalSalesTable].[Region]",
            filterOptions = new
            {
                type = "ValueIsGreaterThan",
                dataFieldName = "[Measures].[TotalRevenue]",
                number1 = 12000
            }
        });
        AssertSuccessful(valueFiltered);
        using (var valueFilter = JsonDocument.Parse(valueFiltered.Result!))
        {
            var filter = Assert.Single(valueFilter.RootElement.GetProperty("filters").EnumerateArray());
            Assert.Equal("ValueIsGreaterThan", filter.GetProperty("type").GetString());
            Assert.Equal(12000, filter.GetProperty("value1").GetInt32());
            Assert.Equal("Revenue", filter.GetProperty("dataFieldName").GetString());
        }
        var valueFilteredRows = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(valueFilteredRows);
        AssertRegionalData(valueFilteredRows.Result!, ("South", 12500d), ("West", 14500d));
        var clearValueFilter = _fixture.Send("pivottablefield.clear-field-filters", new
        {
            pivotTableName = name,
            fieldName = "[RegionalSalesTable].[Region]"
        });
        AssertSuccessful(clearValueFilter);

        var sortedByCaption = _fixture.Send("pivottablefield.sort-field", new
        {
            pivotTableName = name,
            fieldName = "[RegionalSalesTable].[Region]",
            direction = "Descending",
            dataFieldName = "Revenue"
        });
        AssertSuccessful(sortedByCaption);
        var captionSortedRows = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(captionSortedRows);
        AssertRegionalData(captionSortedRows.Result!, ("West", 14500d), ("South", 12500d),
            ("East", 11500d), ("North", 10500d));

        var sortedByUniqueName = _fixture.Send("pivottablefield.sort-field", new
        {
            pivotTableName = name,
            fieldName = "[RegionalSalesTable].[Region]",
            direction = "Ascending",
            dataFieldName = "[Measures].[TotalRevenue]"
        });
        AssertSuccessful(sortedByUniqueName);
        var uniqueNameSortedRows = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(uniqueNameSortedRows);
        AssertRegionalData(uniqueNameSortedRows.Result!, ("North", 10500d), ("East", 11500d),
            ("South", 12500d), ("West", 14500d));
    }

    [Fact]
    public void OlapFieldFilters_AcceptReturnedLevelName()
    {
        var pivot = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"LevelFilter_{Guid.NewGuid():N}";
        RequireSuccess(pivot.CreateFromDataModel(_fixture.BatchToken, "RegionalSalesTable", sheet, "A1", name));
        RequireSuccess(pivot.AddRowField(_fixture.BatchToken, name, "[RegionalSalesTable].[Region]", null));
        RequireSuccess(pivot.AddValueField(_fixture.BatchToken, name, "[Measures].[TotalRevenue]",
            Sbroenne.ExcelMcp.Core.Models.AggregationFunction.Sum, "Revenue"));

        var hierarchyFilters = _fixture.Send("pivottablefield.get-field-filters", new
        {
            pivotTableName = name,
            fieldName = "[RegionalSalesTable].[Region]"
        });
        AssertSuccessful(hierarchyFilters);
        using var hierarchyResult = JsonDocument.Parse(hierarchyFilters.Result!);
        string levelName = hierarchyResult.RootElement.GetProperty("fieldName").GetString()!;
        Assert.Equal("[RegionalSalesTable].[Region].[Region]", levelName);

        var added = _fixture.Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = levelName,
            filterOptions = new { type = "CaptionEquals", text1 = "North" }
        });
        AssertSuccessful(added);
        var filtered = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(filtered);
        AssertRegionalData(filtered.Result!, ("North", 10500d));
        var read = _fixture.Send("pivottablefield.get-field-filters", new
        {
            pivotTableName = name,
            fieldName = levelName
        });
        AssertSuccessful(read);
        using (var result = JsonDocument.Parse(read.Result!))
        {
            Assert.Equal(levelName, result.RootElement.GetProperty("fieldName").GetString());
            var filter = Assert.Single(result.RootElement.GetProperty("filters").EnumerateArray());
            Assert.Equal("CaptionEquals", filter.GetProperty("type").GetString());
            Assert.Equal("North", filter.GetProperty("value1").GetString());
        }
        var layout = _fixture.Send("pivottablecalc.set-layout-options", new
        {
            pivotTableName = name,
            layoutOptions = new { allowMultipleFilters = true }
        });
        AssertSuccessful(layout);
        var combined = _fixture.Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = levelName,
            filterOptions = new { type = "ValueIsGreaterThan", dataFieldName = "Revenue", number1 = 10000 }
        });
        AssertSuccessful(combined);
        using (var result = JsonDocument.Parse(combined.Result!))
        {
            var filters = result.RootElement.GetProperty("filters").EnumerateArray().ToArray();
            Assert.Equal(2, filters.Length);
            Assert.Contains(filters, filter => filter.GetProperty("type").GetString() == "CaptionEquals");
            Assert.Contains(filters, filter => filter.GetProperty("type").GetString() == "ValueIsGreaterThan");
        }
        var combinedData = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(combinedData);
        AssertRegionalData(combinedData.Result!, ("North", 10500d));
        var cleared = _fixture.Send("pivottablefield.clear-field-filters", new
        {
            pivotTableName = name,
            fieldName = levelName
        });
        AssertSuccessful(cleared);
        using (var result = JsonDocument.Parse(cleared.Result!))
            Assert.Empty(result.RootElement.GetProperty("filters").EnumerateArray());
        var restored = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(restored);
        AssertRegionalData(restored.Result!);
    }

    [Fact]
    public async Task OlapCaptionFiltersWorkAndDateFiltersFailWithoutChangingData()
    {
        var pivot = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var regionalName = $"CaptionFilter_{Guid.NewGuid():N}";
        RequireSuccess(pivot.CreateFromDataModel(_fixture.BatchToken, "RegionalSalesTable", sheet, "A1", regionalName));
        RequireSuccess(pivot.AddRowField(_fixture.BatchToken, regionalName, "[RegionalSalesTable].[Region]", null));
        RequireSuccess(pivot.AddValueField(_fixture.BatchToken, regionalName, "[Measures].[TotalRevenue]",
            Sbroenne.ExcelMcp.Core.Models.AggregationFunction.Sum, "Revenue"));

        var captionFilter = _fixture.Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = regionalName,
            fieldName = "[RegionalSalesTable].[Region]",
            filterOptions = new { type = "CaptionEquals", text1 = "North" }
        });
        AssertSuccessful(captionFilter);
        var captionRows = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = regionalName });
        AssertSuccessful(captionRows);
        AssertRegionalData(captionRows.Result!, ("North", 10500d));
        var snapshot = RequireSuccess(pivot.Read(_fixture.BatchToken, regionalName, sheet));
        Assert.True(snapshot.SnapshotComplete);
        var selection = Assert.Single(snapshot.FieldSelections);
        Assert.True(selection.FiltersAvailable);
        var snapshotFilter = Assert.Single(selection.Filters);
        Assert.Equal("CaptionEquals", snapshotFilter.Type);
        Assert.Equal("North", snapshotFilter.Value1);

        var dateName = $"DateFilter_{Guid.NewGuid():N}";
        RequireSuccess(pivot.CreateFromDataModel(_fixture.BatchToken, "SalesTable", sheet, "J1", dateName));
        RequireSuccess(pivot.AddRowField(_fixture.BatchToken, dateName, "[SalesTable].[Date]", null));
        RequireSuccess(pivot.AddValueField(_fixture.BatchToken, dateName, "[Measures].[Total Sales]",
            Sbroenne.ExcelMcp.Core.Models.AggregationFunction.Sum, "Total Sales"));
        var beforeDateFilter = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = dateName });
        AssertSuccessful(beforeDateFilter);
        var dateFilter = await _fixture.SendForFailureAsync("pivottablefield.add-field-filter", new
        {
            pivotTableName = dateName,
            fieldName = "[SalesTable].[Date]",
            filterOptions = new
            {
                type = "DateBetween",
                date1 = new DateTime(2024, 2, 10),
                date2 = new DateTime(2024, 3, 5)
            }
        });
        Assert.False(dateFilter.Success);
        Assert.Equal(nameof(NotSupportedException), dateFilter.ExceptionType);
        Assert.Contains("Excel does not support native date filters on OLAP/Data Model PivotTable fields",
            dateFilter.ErrorMessage, StringComparison.Ordinal);
        var filteredDates = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = dateName });
        AssertSuccessful(filteredDates);
        Assert.Equal(beforeDateFilter.Result, filteredDates.Result);
    }

    [Fact]
    public void OlapLayout_ReadsAllPlacedFieldsAndAppliesNativeStyle()
    {
        var pivot = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Layout_{Guid.NewGuid():N}";
        RequireSuccess(pivot.CreateFromDataModel(_fixture.BatchToken, "RegionalSalesTable", sheet, "A1", name));
        RequireSuccess(pivot.AddRowField(_fixture.BatchToken, name, "[RegionalSalesTable].[Region]", null));
        RequireSuccess(pivot.AddValueField(_fixture.BatchToken, name, "[Measures].[TotalRevenue]",
            Sbroenne.ExcelMcp.Core.Models.AggregationFunction.Sum, "Revenue"));
        var baseline = _fixture.Send("pivottablecalc.set-layout-options", new
        {
            pivotTableName = name,
            layoutOptions = new { rowLayout = 0, repeatLabels = false, styleName = "PivotStyleLight16", preserveFormatting = false }
        });
        AssertSuccessful(baseline);
        using (var original = JsonDocument.Parse(baseline.Result!))
        {
            Assert.False(original.RootElement.GetProperty("preserveFormatting").GetBoolean());
            Assert.Equal(0, Assert.Single(original.RootElement.GetProperty("rowFields").EnumerateArray())
                .GetProperty("rowLayout").GetInt32());
        }
        var before = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(before);
        AssertRegionalData(before.Result!);
        var response = _fixture.Send("pivottablecalc.set-layout-options", new
        {
            pivotTableName = name,
            layoutOptions = new { rowLayout = 1, repeatLabels = true, styleName = "PivotStyleMedium9", preserveFormatting = true }
        });
        AssertSuccessful(response);
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("PivotStyleMedium9", result.RootElement.GetProperty("styleName").GetString());
        Assert.True(result.RootElement.GetProperty("preserveFormatting").GetBoolean());
        var row = Assert.Single(result.RootElement.GetProperty("rowFields").EnumerateArray());
        Assert.Equal("[RegionalSalesTable].[Region].[Region]", row.GetProperty("fieldName").GetString());
        Assert.Equal(1, row.GetProperty("rowLayout").GetInt32());
        Assert.True(row.GetProperty("repeatLabels").GetBoolean());
        var after = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = name });
        AssertSuccessful(after);
        AssertRegionalData(after.Result!);
    }

    private static void AssertSuccessful(Sbroenne.ExcelMcp.Service.ServiceResponse response)
    {
        Assert.True(response.Success, response.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(response.ErrorMessage), response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
    }

    private static void AssertRegionalData(string json)
    {
        AssertRegionalData(json, ("East", 11500d), ("North", 10500d), ("South", 12500d), ("West", 14500d));
    }

    private static void AssertRegionalData(string json, params (string Region, double Revenue)[] expectedRows)
    {
        using var result = JsonDocument.Parse(json);
        var data = result.RootElement.GetProperty("values");
        Assert.Equal(expectedRows.Length + 2, data.GetArrayLength());
        var rows = data.EnumerateArray().Skip(1).Take(expectedRows.Length).ToArray();
        Assert.Equal(expectedRows.Select(expected => expected.Region),
            rows.Select(row => row[0].GetString()));
        Assert.Equal(expectedRows.Select(expected => expected.Revenue),
            rows.Select(row => row[1].GetDouble()));
        Assert.Equal(expectedRows.Sum(expected => expected.Revenue), data[data.GetArrayLength() - 1][1].GetDouble());
    }
}
