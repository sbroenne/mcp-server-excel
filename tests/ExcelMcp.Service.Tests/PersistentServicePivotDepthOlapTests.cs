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
        Assert.Equal(before.Result, after.Result);
    }

    [Fact]
    public void OlapLayout_ReadsAllPlacedFieldsAndAppliesNativeStyle()
    {
        var response = _fixture.Send("pivottablecalc.set-layout-options", new
        {
            pivotTableName = "DataModelPivot",
            layoutOptions = new { rowLayout = 1, repeatLabels = true, styleName = "PivotStyleMedium9", preserveFormatting = true }
        });
        using var result = System.Text.Json.JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal("PivotStyleMedium9", result.RootElement.GetProperty("styleName").GetString());
        Assert.NotEmpty(result.RootElement.GetProperty("rowFields").EnumerateArray());
    }
}
