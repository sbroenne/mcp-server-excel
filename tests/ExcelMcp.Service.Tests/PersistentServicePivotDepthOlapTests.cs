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
    private static readonly string[] ExpectedRows = ["East", "North", "South", "West", "Grand Total"];
    private static readonly double[] ExpectedTotals = [11500, 10500, 12500, 14500, 49000];
    [Theory]
    [InlineData("pivottablefield.get-field-filters")]
    [InlineData("pivottablefield.clear-field-filters")]
    [InlineData("pivottablefield.add-field-filter")]
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
        if (action.EndsWith("add-field-filter", StringComparison.Ordinal))
            args["filterOptions"] = new { type = "CaptionEquals", text1 = "North" };
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
        Assert.Contains("OLAP", failure.ErrorMessage, StringComparison.Ordinal);
        var after = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = "DataModelPivot" });
        AssertSuccessful(after);
        Assert.Equal(before.Result, after.Result);
        AssertRegionalData(after.Result!);
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
        using var result = JsonDocument.Parse(json);
        var data = result.RootElement.GetProperty("values");
        Assert.Equal(6, data.GetArrayLength());
        Assert.Equal(ExpectedRows,
            data.EnumerateArray().Skip(1).Select(row => row[0].GetString()));
        Assert.Equal(ExpectedTotals,
            data.EnumerateArray().Skip(1).Select(row => row[1].GetDouble()));
    }
}
