using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceDataModelDeleteTableTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IDataModelCommands _dataModel =
        fixture.CreateCommands<IDataModelCommands>();
    private readonly ITableCommands _tables =
        fixture.CreateCommands<ITableCommands>();

    [Fact]
    public void DeleteTable_ExistingTable_RemovesFromDataModel()
    {
        CreateDataModelTable();

        var listBefore = _dataModel.ListTables(_fixture.BatchToken);
        Assert.True(listBefore.Success);
        Assert.Contains(listBefore.Tables, table => table.Name == "TestTable");

        _dataModel.DeleteTable(_fixture.BatchToken, "TestTable");

        var listAfter = _dataModel.ListTables(_fixture.BatchToken);
        Assert.True(listAfter.Success);
        Assert.DoesNotContain(listAfter.Tables, table => table.Name == "TestTable");
    }

    [Fact]
    public void DeleteTable_NonExistentTable_ThrowsInvalidOperationException()
    {
        CreateDataModelTable();
        _fixture.RegisterDataModelTableForCleanup("TestTable");

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _dataModel.DeleteTable(_fixture.BatchToken, "NonExistentTable"));

        Assert.Contains(
            "not found",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private void CreateDataModelTable()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _commands.SetValues(
            _fixture.BatchToken,
            sheetName,
            "A1:B3",
            [["ID", "Value"], [1, 100], [2, 200]]);
        _tables.Create(
            _fixture.BatchToken,
            sheetName,
            "TestTable",
            "A1:B3");
        _tables.AddToDataModel(_fixture.BatchToken, "TestTable");
    }
}
