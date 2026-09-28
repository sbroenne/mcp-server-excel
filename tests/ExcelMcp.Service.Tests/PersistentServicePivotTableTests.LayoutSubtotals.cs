using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SetLayout_RoundTrip_PersistsLayoutChange()
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", destinationSheet, "A1", "SalesPivot");
        Assert.True(createResult.Success);

        var row1 = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        Assert.True(row1.Success);

        var row2 = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        Assert.True(row2.Success);

        var value = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        Assert.True(value.Success);

        var layoutResult = _pivotCommands.SetLayout(batch, "SalesPivot", 1);
        Assert.True(layoutResult.Success);

        await _fixture.SaveAndReopenAsync();

        batch = _fixture.BatchToken;
        var listResult = _pivotCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.PivotTables, pt => pt.Name == "SalesPivot");

        var fields = _pivotCommands.ListFields(batch, "SalesPivot");
        Assert.True(fields.Success);
        Assert.Contains(fields.Fields, f => f.Name == "Region");
        Assert.Contains(fields.Fields, f => f.Name == "Product");
    }
}
