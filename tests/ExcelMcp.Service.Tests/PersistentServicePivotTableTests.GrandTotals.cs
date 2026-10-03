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
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Product"));
        RequireSuccess(_pivotCommands.AddColumnField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));
        RequireSuccess(_pivotCommands.SetGrandTotals(batch, pivotName, !showRowGrandTotals, !showColumnGrandTotals));
        Assert.Equal((!showRowGrandTotals, !showColumnGrandTotals), ReadGrandTotals(destinationSheet, pivotName));

        var setResult = _pivotCommands.SetGrandTotals(
            batch,
            pivotName,
            showRowGrandTotals,
            showColumnGrandTotals);
        RequireSuccess(setResult);

        Assert.Equal(
            (showRowGrandTotals, showColumnGrandTotals),
            ReadGrandTotals(destinationSheet, pivotName));
        AssertGrandTotalCells(pivotName, showRowGrandTotals, showColumnGrandTotals);
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
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, "TestPivot", "Product"));
        RequireSuccess(_pivotCommands.AddColumnField(batch, "TestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, "TestPivot", "Sales"));

        var setResult = _pivotCommands.SetGrandTotals(
                batch,
                "TestPivot",
                true,
                false);
        RequireSuccess(setResult);

        await _fixture.SaveAndReopenAsync();

        var readResult = _pivotCommands.Read(_fixture.BatchToken, "TestPivot");
        RequireSuccess(readResult);
        Assert.Equal(
            (true, false),
            ReadGrandTotals(destinationSheet, "TestPivot"));
        RequireSuccess(readResult);
        AssertGrandTotalCells("TestPivot", true, false);
    }

    private (bool Row, bool Column) ReadGrandTotals(
        string sheetName,
        string pivotName) =>
        ReadNativePivot(sheetName, pivotName, pivot => (pivot.RowGrand, pivot.ColumnGrand));

    private void AssertGrandTotalCells(string pivotName, bool rowGrand, bool columnGrand)
    {
        var values = RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, pivotName)).Values;
        Assert.Equal(columnGrand ? 5 : 4, values.Count);
        Assert.All(values, row => Assert.Equal(rowGrand ? 4 : 3, row.Count));
        Assert.Equal("North", values[1][1]);
        Assert.Equal("South", values[1][2]);
        Assert.Equal("Gadget", values[2][0]);
        Assert.Equal("Widget", values[3][0]);
        Assert.Equal(75d, Convert.ToDouble(values[2][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(200d, Convert.ToDouble(values[2][2], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(250d, Convert.ToDouble(values[3][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(125d, Convert.ToDouble(values[3][2], System.Globalization.CultureInfo.InvariantCulture));
        if (rowGrand)
        {
            Assert.Equal(275d, Convert.ToDouble(values[2][3], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(375d, Convert.ToDouble(values[3][3], System.Globalization.CultureInfo.InvariantCulture));
        }
        if (columnGrand)
        {
            Assert.Equal(325d, Convert.ToDouble(values[4][1], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(325d, Convert.ToDouble(values[4][2], System.Globalization.CultureInfo.InvariantCulture));
            if (rowGrand)
            {
                Assert.Equal(650d, Convert.ToDouble(values[4][3], System.Globalization.CultureInfo.InvariantCulture));
            }
        }
        AssertOriginalSales();
    }
}
