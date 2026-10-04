using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceDataModelNoModelTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IDataModelCommands _dataModel =
        fixture.CreateCommands<IDataModelCommands>();
    private string _guardSheetName = string.Empty;

    [Fact]
    public void Refresh_NoDataModel_ThrowsInvalidOperationException()
    {
        SeedGuard();
        var exception = Assert.Throws<InvalidOperationException>(
            () => _dataModel.Refresh(_fixture.BatchToken));
        Assert.Contains("Data Model", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertGuardPreserved();
    }

    [Fact]
    public void DeleteTable_EmptyDataModel_ThrowsInvalidOperationException()
    {
        SeedGuard();
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _dataModel.DeleteTable(_fixture.BatchToken, "AnyTable"));

        Assert.Contains(
            "no tables",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        AssertGuardPreserved();
    }

    [Fact]
    public void RenameTable_EmptyDataModel_ReturnsFailure()
    {
        SeedGuard();
        var result = _dataModel.RenameTable(
            _fixture.BatchToken,
            "AnyTable",
            "NewName");

        Assert.False(result.Success);
        Assert.Contains(
            "no tables",
            result.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
        AssertGuardPreserved();
    }

    [Fact]
    public async Task Evaluate_MissingModel_HasPrerequisiteCategory()
    {
        SeedGuard();
        var response = await _fixture.SendForFailureAsync(
            "datamodel.evaluate",
            new { daxQuery = "EVALUATE ROW(\"Value\", 1)" });

        Assert.Contains(
            "Data Model",
            response.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal(
            OperationFailureCategory.Prerequisite.ToString(),
            response.ErrorCategory);
        AssertGuardPreserved();
    }

    private void SeedGuard()
    {
        _guardSheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, _guardSheetName, "A1:B1", [["retained", 7]]));
        RequireSuccess(_commands.SetFormulas(_fixture.BatchToken, _guardSheetName, "C1", [["=B1*2"]]));
        AssertGuardPreserved();
    }

    private void AssertGuardPreserved()
    {
        Assert.Empty(RequireSuccess(_dataModel.ListTables(_fixture.BatchToken)).Tables);
        var cells = RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, _guardSheetName, "A1:C1"));
        Assert.Equal(["retained", 7, 14], Assert.Single(cells.Values));
        Assert.Equal("=B1*2", Assert.Single(cells.Formulas)[2]);
    }
}
