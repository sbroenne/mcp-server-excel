using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SetLayout_AcceptsEverySupportedLayout(int layout)
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        var pivotName = $"Layout_{Guid.NewGuid():N}";
        var createResult = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName,
            "A1:D6",
            destinationSheet,
            "A1",
            pivotName);
        Assert.True(createResult.Success, createResult.ErrorMessage);
        Assert.True(
            _pivotCommands.AddRowField(batch, pivotName, "Region").Success);
        Assert.True(
            _pivotCommands.AddRowField(batch, pivotName, "Product").Success);
        Assert.True(
            _pivotCommands.AddValueField(batch, pivotName, "Sales").Success);

        var result = _pivotCommands.SetLayout(batch, pivotName, layout);

        Assert.True(result.Success, result.ErrorMessage);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SetSubtotals_AppliesVisibility(bool show)
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        var pivotName = $"Subtotals_{Guid.NewGuid():N}";
        var createResult = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName,
            "A1:D6",
            destinationSheet,
            "A1",
            pivotName);
        Assert.True(createResult.Success, createResult.ErrorMessage);
        Assert.True(
            _pivotCommands.AddRowField(batch, pivotName, "Region").Success);
        Assert.True(
            _pivotCommands.AddValueField(batch, pivotName, "Sales").Success);

        var result = _pivotCommands.SetSubtotals(
            batch,
            pivotName,
            "Region",
            show);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal("Region", result.FieldName);
    }

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
