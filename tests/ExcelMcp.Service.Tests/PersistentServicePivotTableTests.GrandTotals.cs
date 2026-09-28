using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Fact]
    public async Task SetGrandTotals_RoundTrip_PersistsConfiguration()
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        var createResult = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName,
            "A1:D6",
            destinationSheet,
            "A1",
            "TestPivot");
        Assert.True(createResult.Success, createResult.ErrorMessage);
        Assert.True(_pivotCommands.AddRowField(batch, "TestPivot", "Product").Success);
        Assert.True(_pivotCommands.AddColumnField(batch, "TestPivot", "Region").Success);
        Assert.True(_pivotCommands.AddValueField(batch, "TestPivot", "Sales").Success);

        var setResult = _pivotCommands.SetGrandTotals(
                batch,
                "TestPivot",
                true,
                false);
        Assert.True(
            setResult.Success,
            $"SetGrandTotals failed: {setResult.ErrorMessage}");

        await _fixture.SaveAndReopenAsync();

        var readResult = _pivotCommands.Read(_fixture.BatchToken, "TestPivot");
        Assert.True(
            readResult.Success,
            $"Read failed: {readResult.ErrorMessage}");
    }
}
