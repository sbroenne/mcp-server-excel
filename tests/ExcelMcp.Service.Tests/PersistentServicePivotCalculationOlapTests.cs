using System.Text.Json;
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
    private readonly IPersistentPivotTableCommands _pivot =
        fixture.CreateCommands<IPersistentPivotTableCommands>();

    [Fact]
    public async Task DataModel_IgnoredNativeCalculation_FailsInsteadOfClaimingSuccess()
    {
        var before = _pivot.ListFields(_fixture.BatchToken, "DataModelPivot");
        Assert.True(before.Success, before.ErrorMessage);
        var value = Assert.Single(before.ValueFields);
        Assert.True(value.IsOlap);
        Assert.False(value.BaseSettingsAvailable);
        Assert.Null(value.Function);
        var failure = await Record.ExceptionAsync(async () =>
        {
            var rejected = await _fixture.SendForFailureAsync("pivottablefield.set-field-calculation", new
            {
                pivotTableName = "DataModelPivot",
                fieldName = value.FieldName,
                calculation = "PercentOfTotal"
            });
            Assert.False(rejected.Success);
            Assert.Contains("did not apply", rejected.ErrorMessage);
            Assert.Contains("source measure", rejected.ErrorMessage);
            var listed = _pivot.ListFields(_fixture.BatchToken, "DataModelPivot");
            Assert.True(listed.Success, listed.ErrorMessage);
            var read = Assert.Single(listed.ValueFields);
            Assert.Equal(value.SourceName, read.SourceName);
            Assert.Equal(value.Calculation, read.Calculation);
        });
        var restoration = Record.Exception(() =>
        {
            var restore = _pivot.SetFieldCalculation(_fixture.BatchToken, "DataModelPivot",
                value.FieldName, value.Calculation);
            Assert.True(restore.Success, restore.ErrorMessage);
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
        var response = await _fixture.SendForFailureAsync("pivottablefield.set-field-calculation", new
        {
            pivotTableName = "DataModelPivot",
            fieldName = value.FieldName,
            calculation = "RunningTotal",
            baseFieldName = "[RegionalSalesTable].[Region]"
        });
        Assert.False(response.Success);
        Assert.Contains("BaseField/BaseItem", response.ErrorMessage);
        var after = _fixture.Send("pivottablefield.list-fields", new { pivotTableName = "DataModelPivot" });
        using var read = JsonDocument.Parse(after.Result!);
        Assert.Equal(value.Calculation.ToString(), Assert.Single(read.RootElement.GetProperty("valueFields").EnumerateArray())
            .GetProperty("calculation").GetString());
    }
}
