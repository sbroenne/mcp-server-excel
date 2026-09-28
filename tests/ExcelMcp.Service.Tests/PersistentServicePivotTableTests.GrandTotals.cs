using Sbroenne.ExcelMcp.ComInterop;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void SetGrandTotals_AppliesRowAndColumnConfiguration(
        bool showRowGrandTotals,
        bool showColumnGrandTotals)
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        var pivotName = $"Totals_{Guid.NewGuid():N}";
        var createResult = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName,
            "A1:D6",
            destinationSheet,
            "A1",
            pivotName);
        Assert.True(createResult.Success, createResult.ErrorMessage);

        var setResult = _pivotCommands.SetGrandTotals(
            batch,
            pivotName,
            showRowGrandTotals,
            showColumnGrandTotals);
        Assert.True(setResult.Success, setResult.ErrorMessage);

        Assert.Equal(
            (showRowGrandTotals, showColumnGrandTotals),
            ReadGrandTotals(destinationSheet, pivotName));
    }

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
        Assert.Equal(
            (true, false),
            ReadGrandTotals(destinationSheet, "TestPivot"));
    }

    private (bool Row, bool Column) ReadGrandTotals(
        string sheetName,
        string pivotName) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            dynamic? sheet = null;
            dynamic? pivot = null;
            try
            {
                sheet = context.Book.Worksheets.Item[sheetName];
                pivot = sheet.PivotTables(pivotName);
                return ((bool)pivot.RowGrand, (bool)pivot.ColumnGrand);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
                ComUtilities.Release(ref sheet);
            }
        });
}
