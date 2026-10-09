using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePivotModelConnectionTests(PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture), IClassFixture<PersistentServiceDataModelFixture>
{
    [Fact]
    public async Task SetConnection_WorkbookDataModelIsReportedAndCannotBeRedirected()
    {
        var commands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        var sheet = RequireSuccess(commands.List(_fixture.BatchToken)).PivotTables
            .Single(pivot => pivot.Name == "DataModelPivot").SheetName;
        var before = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = "DataModelPivot" });
        Assert.True(before.Success, before.ErrorMessage);
        var read = _fixture.Send("pivottable.get-connection",
            new { sheetName = sheet, pivotTableName = "DataModelPivot" });
        Assert.True(read.Success, read.ErrorMessage);
        using var source = JsonDocument.Parse(read.Result!);
        Assert.True(source.RootElement.GetProperty("isDataModel").GetBoolean());
        Assert.True(source.RootElement.GetProperty("isOlap").GetBoolean());
        Assert.False(source.RootElement.GetProperty("isExternal").GetBoolean());
        var failure = await _fixture.SendForFailureAsync("pivottable.set-connection",
            new { sheetName = sheet, pivotTableName = "DataModelPivot", connectionName = "MissingConnection" });
        Assert.Contains("Data Model", failure.ErrorMessage, StringComparison.Ordinal);
        var after = _fixture.Send("pivottablecalc.get-data", new { pivotTableName = "DataModelPivot" });
        Assert.True(after.Success, after.ErrorMessage);
        Assert.Equal(before.Result, after.Result);
        var unchanged = _fixture.Send("pivottable.get-connection",
            new { sheetName = sheet, pivotTableName = "DataModelPivot" });
        Assert.True(unchanged.Success, unchanged.ErrorMessage);
        Assert.Equal(read.Result, unchanged.Result);
    }
}
