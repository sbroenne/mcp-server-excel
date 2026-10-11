using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotCalculation")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePivotCalculationOlapTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private static readonly string[] ExpectedRows = ["East", "North", "South", "West", "Grand Total"];
    private static readonly double[] ExpectedTotals = [11500, 10500, 12500, 14500, 49000];
    private readonly IPersistentPivotTableCommands _pivot =
        fixture.CreateCommands<IPersistentPivotTableCommands>();

    [Fact]
    public async Task DataModel_StaleNativeCalculationReadback_ReportsFailureAndChangedDisplayedValues()
    {
        var before = _pivot.ListFields(_fixture.BatchToken, "DataModelPivot");
        Assert.True(before.Success, before.ErrorMessage);
        var value = Assert.Single(before.ValueFields);
        Assert.True(value.IsOlap);
        Assert.False(value.BaseSettingsAvailable);
        Assert.Null(value.Function);
        Assert.Equal(PivotFieldCalculation.Normal, value.Calculation);
        AssertRegionalData();
        var failure = await Record.ExceptionAsync(async () =>
        {
            var rejected = await _fixture.SendForFailureAsync("pivottablefield.set-field-calculation", new
            {
                pivotTableName = "DataModelPivot",
                fieldName = value.FieldName,
                calculation = "PercentOfTotal"
            });
            Assert.False(rejected.Success);
            Assert.Equal(nameof(InvalidOperationException), rejected.ExceptionType);
            Assert.Contains($"Excel accepted the Show Values As write for '{value.FieldName}'", rejected.ErrorMessage);
            Assert.Contains("native readback did not verify", rejected.ErrorMessage);
            Assert.Contains("Requested calculation: PercentOfTotal; native calculation: Normal", rejected.ErrorMessage);
            Assert.Contains("Displayed values may already have changed", rejected.ErrorMessage);
            Assert.Contains("Inspect get-data", rejected.ErrorMessage);
            Assert.Contains("no changes have been rolled back", rejected.ErrorMessage);
            Assert.DoesNotContain("did not apply", rejected.ErrorMessage);
            var listed = _pivot.ListFields(_fixture.BatchToken, "DataModelPivot");
            Assert.True(listed.Success, listed.ErrorMessage);
            var read = Assert.Single(listed.ValueFields);
            Assert.Equal(value.SourceName, read.SourceName);
            Assert.Equal(value.FieldName, read.FieldName);
            Assert.Equal(value.Calculation, read.Calculation);
            AssertRegionalData(ExpectedTotals.Select(total => total / 49000).ToArray());
        });
        var restoration = Record.Exception(() =>
        {
            var restore = _pivot.SetFieldCalculation(_fixture.BatchToken, "DataModelPivot",
                value.FieldName, PivotFieldCalculation.Normal);
            Assert.True(restore.Success, restore.ErrorMessage);
            AssertRegionalData();
        });
        if (restoration is not null)
            failure = PersistentServiceCleanupFailures.Combine(failure, restoration);
        if (failure is not null)
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
    }

    [Fact]
    public async Task DataModel_BaseDependentCalculation_IsRejectedBeforeChangingMeasure()
    {
        var before = _pivot.ListFields(_fixture.BatchToken, "DataModelPivot");
        Assert.True(before.Success, before.ErrorMessage);
        var value = Assert.Single(before.ValueFields);
        var dataBefore = AssertRegionalData();
        var response = await _fixture.SendForFailureAsync("pivottablefield.set-field-calculation", new
        {
            pivotTableName = "DataModelPivot",
            fieldName = value.FieldName,
            calculation = "RunningTotal",
            baseFieldName = "[RegionalSalesTable].[Region]"
        });
        Assert.False(response.Success);
        Assert.Equal(nameof(NotSupportedException), response.ExceptionType);
        Assert.Contains("Excel does not expose BaseField/BaseItem for OLAP/Data Model PivotTables", response.ErrorMessage);
        var after = _fixture.Send("pivottablefield.list-fields", new { pivotTableName = "DataModelPivot" });
        using var read = JsonDocument.Parse(after.Result!);
        Assert.Equal(value.Calculation.ToString(), Assert.Single(read.RootElement.GetProperty("valueFields").EnumerateArray())
            .GetProperty("calculation").GetString());
        Assert.Equal(dataBefore, AssertRegionalData());
    }

    private string AssertRegionalData(double[]? expectedValues = null)
    {
        var response = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = "DataModelPivot" });
        Assert.True(response.Success, response.ErrorMessage);
        using var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        var values = result.RootElement.GetProperty("values");
        Assert.Equal(6, values.GetArrayLength());
        Assert.Equal(ExpectedRows, values.EnumerateArray().Skip(1).Select(row => row[0].GetString()));
        Assert.Equal(expectedValues ?? ExpectedTotals, values.EnumerateArray().Skip(1).Select(row => row[1].GetDouble()));
        return response.Result!;
    }
}
