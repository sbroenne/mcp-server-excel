using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    [Theory]
    [InlineData(-1)]
    [InlineData(3)]
    public void SetLayout_InvalidLayout_PreservesConfiguredPivot(int layout)
    {
        RequireSuccess(_pivotCommands.CreateFromRange(
            _fixture.BatchToken, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot"));
        RequireSuccess(_pivotCommands.AddRowField(_fixture.BatchToken, "TestPivot", "Region"));
        RequireSuccess(_pivotCommands.AddValueField(_fixture.BatchToken, "TestPivot", "Sales"));
        RequireSuccess(_pivotCommands.SetLayout(_fixture.BatchToken, "TestPivot", 1));
        AssertPivotSales(325, 325);
        var before = SnapshotPivot();
        var data = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, "TestPivot")));

        var error = Assert.Throws<ArgumentException>(() =>
            _pivotCommands.SetLayout(_fixture.BatchToken, "TestPivot", layout));

        Assert.Contains("pivottablecalc.set-layout failed [InvalidInput/ArgumentException]",
            error.Message, StringComparison.Ordinal);
        Assert.Equal(before, SnapshotPivot());
        Assert.Equal(data, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_pivotCommands.GetData(_fixture.BatchToken, "TestPivot"))));
        AssertNativeLayout(_salesSheetName, "TestPivot", 1);
    }

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
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Product"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));

        var result = _pivotCommands.SetLayout(batch, pivotName, layout);

        RequireSuccess(result);
        AssertNativeLayout(destinationSheet, pivotName, layout);
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
        RequireSuccess(createResult);
        RequireSuccess(_pivotCommands.AddRowField(batch, pivotName, "Region"));
        RequireSuccess(_pivotCommands.AddValueField(batch, pivotName, "Sales"));

        var result = _pivotCommands.SetSubtotals(
            batch,
            pivotName,
            "Region",
            show);

        RequireSuccess(result);
        Assert.Equal("Region", result.FieldName);
        RequireSuccess(result);
        ReadNativePivot(destinationSheet, pivotName, pivot =>
        {
            Microsoft.Office.Interop.Excel.PivotField? field = null;
            try
            {
                field = (Microsoft.Office.Interop.Excel.PivotField)pivot.PivotFields("Region");
                for (var index = 1; index <= 12; index++)
                {
                    Assert.Equal(index == 1 && show, Convert.ToBoolean(field.Subtotals[index]));
                }
                return 0;
            }
            finally
            {
                Sbroenne.ExcelMcp.ComInterop.ComUtilities.Release(ref field);
            }
        });
        AssertOriginalSales();
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SetLayout_RoundTrip_PersistsLayoutChange()
    {
        var batch = _fixture.BatchToken;
        var destinationSheet = _fixture.CreateTestSheet(batch);
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", destinationSheet, "A1", "SalesPivot");
        RequireSuccess(createResult);

        var row1 = _pivotCommands.AddRowField(batch, "SalesPivot", "Region");
        RequireSuccess(row1);

        var row2 = _pivotCommands.AddRowField(batch, "SalesPivot", "Product");
        RequireSuccess(row2);

        var value = _pivotCommands.AddValueField(batch, "SalesPivot", "Sales");
        RequireSuccess(value);

        var layoutResult = _pivotCommands.SetLayout(batch, "SalesPivot", 1);
        RequireSuccess(layoutResult);
        AssertNativeLayout(destinationSheet, "SalesPivot", 1);

        await _fixture.SaveAndReopenAsync();

        batch = _fixture.BatchToken;
        var listResult = _pivotCommands.List(batch);
        RequireSuccess(listResult);
        Assert.Contains(listResult.PivotTables, pt => pt.Name == "SalesPivot");

        var fields = _pivotCommands.ListFields(batch, "SalesPivot");
        RequireSuccess(fields);
        Assert.Contains(fields.Fields, f => f.Name == "Region");
        Assert.Contains(fields.Fields, f => f.Name == "Product");
        RequireSuccess(fields);
        AssertNativeLayout(destinationSheet, "SalesPivot", 1);
        var data = RequireSuccess(_pivotCommands.GetData(batch, "SalesPivot"));
        Assert.Equal(650d, Convert.ToDouble(data.Values[^1][^1],
            System.Globalization.CultureInfo.InvariantCulture));
    }
}
