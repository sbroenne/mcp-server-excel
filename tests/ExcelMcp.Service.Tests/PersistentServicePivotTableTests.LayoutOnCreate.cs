using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task CreateFromRange_LayoutAndFields_PersistsAfterSaveAndReopen()
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", destinationSheet, "A1", "PersistPivot");
        RequireSuccess(createResult);

        var layoutResult = _pivotCommands.SetLayout(batch, "PersistPivot", 1);
        RequireSuccess(layoutResult);

        var row = _pivotCommands.AddRowField(batch, "PersistPivot", "Region");
        RequireSuccess(row);

        var value = _pivotCommands.AddValueField(batch, "PersistPivot", "Sales");
        RequireSuccess(value);
        AssertNativeLayout(destinationSheet, "PersistPivot", 1);
        AssertPivotSales(325, 325, "PersistPivot");

        await _fixture.SaveAndReopenAsync();

        batch = _fixture.BatchToken;
        var listResult = _pivotCommands.List(batch);
        RequireSuccess(listResult);
        Assert.Contains(listResult.PivotTables, pt => pt.Name == "PersistPivot");

        var fields = _pivotCommands.ListFields(batch, "PersistPivot");
        RequireSuccess(fields);
        Assert.Contains(fields.Fields, f => f.Name == "Region");
        RequireSuccess(fields);
        AssertNativeLayout(destinationSheet, "PersistPivot", 1);
        AssertNativeField("PersistPivot", "Region", Sbroenne.ExcelMcp.Core.Models.PivotFieldArea.Row, destinationSheet);
        AssertNativeField("PersistPivot", "Sales", Sbroenne.ExcelMcp.Core.Models.PivotFieldArea.Value,
            destinationSheet, Sbroenne.ExcelMcp.Core.Models.AggregationFunction.Sum);
        AssertPivotSales(325, 325, "PersistPivot");
    }
}
