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
        var sheetName = CreateDataModelTable();

        var listBefore = _dataModel.ListTables(_fixture.BatchToken);
        Assert.True(listBefore.Success);
        Assert.Contains(listBefore.Tables, table => table.Name == "TestTable");

        RequireSuccess(_dataModel.DeleteTable(_fixture.BatchToken, "TestTable"));

        var listAfter = _dataModel.ListTables(_fixture.BatchToken);
        Assert.True(listAfter.Success);
        Assert.DoesNotContain(listAfter.Tables, table => table.Name == "TestTable");
        Assert.Empty(listAfter.Tables);
        AssertSourcePreserved(sheetName);
    }

    [Fact]
    public void DeleteTable_NonExistentTable_ThrowsInvalidOperationException()
    {
        var sheetName = CreateDataModelTable();
        _fixture.RegisterDataModelTableForCleanup("TestTable");
        var before = RequireSuccess(_dataModel.ReadTable(_fixture.BatchToken, "TestTable"));

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _dataModel.DeleteTable(_fixture.BatchToken, "NonExistentTable"));

        Assert.Contains(
            "not found",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
        var after = RequireSuccess(_dataModel.ReadTable(_fixture.BatchToken, "TestTable"));
        Assert.Equal(System.Text.Json.JsonSerializer.Serialize(before),
            System.Text.Json.JsonSerializer.Serialize(after));
        Assert.Equal("TestTable", Assert.Single(
            RequireSuccess(_dataModel.ListTables(_fixture.BatchToken)).Tables).Name);
        AssertSourcePreserved(sheetName);
        AssertModelRows();
    }

    private string CreateDataModelTable()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(
            _fixture.BatchToken,
            sheetName,
            "A1:B3",
            [["ID", "Value"], [1, 100], [2, 200]]));
        RequireSuccess(_tables.Create(
            _fixture.BatchToken,
            sheetName,
            "TestTable",
            "A1:B3"));
        RequireSuccess(_tables.AddToDataModel(_fixture.BatchToken, "TestTable"));
        AssertSourcePreserved(sheetName);
        var table = Assert.Single(RequireSuccess(_dataModel.ListTables(_fixture.BatchToken)).Tables);
        Assert.Equal("TestTable", table.Name);
        Assert.Equal(2, table.RecordCount);
        AssertModelRows();
        return sheetName;
    }

    private void AssertModelRows()
    {
        var result = RequireSuccess(_dataModel.Evaluate(_fixture.BatchToken,
            "EVALUATE 'TestTable' ORDER BY 'TestTable'[ID]"));
        Assert.Equal(["TestTable[ID]", "TestTable[Value]"], result.Columns);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal([1m, 100m], result.Rows[0].Select(value =>
            Convert.ToDecimal(value, System.Globalization.CultureInfo.InvariantCulture)));
        Assert.Equal([2m, 200m], result.Rows[1].Select(value =>
            Convert.ToDecimal(value, System.Globalization.CultureInfo.InvariantCulture)));
    }

    private void AssertSourcePreserved(string sheetName)
    {
        var source = RequireSuccess(_tables.GetData(_fixture.BatchToken, "TestTable"));
        Assert.Equal(["ID", "Value"], source.Headers);
        Assert.Equal(2, source.Data.Count);
        Assert.Equal([1, 100], source.Data[0]);
        Assert.Equal([2, 200], source.Data[1]);
        Assert.Equal(sheetName, RequireSuccess(_tables.Read(_fixture.BatchToken, "TestTable")).Table!.SheetName);
    }
}
