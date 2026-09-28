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

    [Fact]
    public void Refresh_NoDataModel_ThrowsInvalidOperationException()
    {
        Assert.Throws<InvalidOperationException>(
            () => _dataModel.Refresh(_fixture.BatchToken));
    }

    [Fact]
    public void DeleteTable_EmptyDataModel_ThrowsInvalidOperationException()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _dataModel.DeleteTable(_fixture.BatchToken, "AnyTable"));

        Assert.Contains(
            "no tables",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RenameTable_EmptyDataModel_ReturnsFailure()
    {
        var result = _dataModel.RenameTable(
            _fixture.BatchToken,
            "AnyTable",
            "NewName");

        Assert.False(result.Success);
        Assert.Contains(
            "no tables",
            result.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Evaluate_MissingModel_HasPrerequisiteCategory()
    {
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
    }
}
