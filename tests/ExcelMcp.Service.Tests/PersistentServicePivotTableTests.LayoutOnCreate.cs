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
        Assert.True(createResult.Success);

        var layoutResult = _pivotCommands.SetLayout(batch, "PersistPivot", 1);
        Assert.True(layoutResult.Success);

        var row = _pivotCommands.AddRowField(batch, "PersistPivot", "Region");
        Assert.True(row.Success);

        var value = _pivotCommands.AddValueField(batch, "PersistPivot", "Sales");
        Assert.True(value.Success);

        await _fixture.SaveAndReopenAsync();

        batch = _fixture.BatchToken;
        var listResult = _pivotCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.PivotTables, pt => pt.Name == "PersistPivot");

        var fields = _pivotCommands.ListFields(batch, "PersistPivot");
        Assert.True(fields.Success);
        Assert.Contains(fields.Fields, f => f.Name == "Region");
    }
}
