using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for TableCommands.AddToDataModel, focusing on bracket column name detection
/// and stripping. Regression tests for the stripBracketColumnNames feature.
/// </summary>
[Collection("ServiceWorkflow")]
public class PersistentServiceTableDataModelTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly ITableCommands _tableCommands =
        fixture.CreateCommands<ITableCommands>();
    private readonly IDataModelCommands _model =
        fixture.CreateCommands<IDataModelCommands>();

    private void CreateTableWithBracketColumns()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        RequireSuccess(_commands.SetValues(
            batch,
            "Data",
            "A1:C3",
            [
                ["ProductName", "[ACR_CM1]", "[ACR_CM2]"],
                ["Widget", 100.0, 200.0],
                ["Gadget", 150.0, 250.0]
            ]));
        var tableResult = _tableCommands.Create(
            batch,
            "Data",
            "BracketTable",
            "A1:C3");
        Assert.True(tableResult.Success, $"Setup failed: {tableResult.ErrorMessage}");
        _fixture.RegisterTableForCleanup("BracketTable");
    }

    private void CreateTableWithNormalColumns()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        RequireSuccess(_commands.SetValues(
            batch,
            "Data",
            "A1:B2",
            [["ProductName", "Amount"], ["Widget", 100.0]]));
        var tableResult = _tableCommands.Create(
            batch,
            "Data",
            "NormalTable",
            "A1:B2");
        Assert.True(tableResult.Success, $"Setup failed: {tableResult.ErrorMessage}");
        _fixture.RegisterTableForCleanup("NormalTable");
    }

    private AddToDataModelResult AddToDataModel(
        string tableName,
        bool stripBracketColumnNames = false)
    {
        var result = _tableCommands.AddToDataModel(
            _fixture.BatchToken,
            tableName,
            stripBracketColumnNames);
        _fixture.RegisterDataModelTableForCleanup(tableName);
        return result;
    }

    /// <summary>
    /// When a table has bracket column names and stripBracketColumnNames=false,
    /// BracketColumnsFound should be populated with the bracket column names.
    /// </summary>
    [Fact]
    [Trait("Layer", "Service")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Table")]
    [Trait("Feature", "DataModel")]
    [Trait("RequiresExcel", "true")]
    [Trait("Speed", "Medium")]
    public void AddToDataModel_BracketColumns_WithoutStrip_ReturnsBracketColumnsFound()
    {
        // Arrange
        CreateTableWithBracketColumns();

        // Act
        var result = AddToDataModel("BracketTable", stripBracketColumnNames: false);

        // Assert
        Assert.True(result.Success, $"AddToDataModel failed: {result.ErrorMessage}");
        Assert.NotNull(result.BracketColumnsFound);
        Assert.Equal(2, result.BracketColumnsFound.Length);
        Assert.Contains("[ACR_CM1]", result.BracketColumnsFound);
        Assert.Contains("[ACR_CM2]", result.BracketColumnsFound);
        Assert.Null(result.BracketColumnsRenamed);
        AssertLoadedTable("BracketTable", ["ProductName", "[ACR_CM1]", "[ACR_CM2]"],
            [["Widget", 100, 200], ["Gadget", 150, 250]]);
    }

    /// <summary>
    /// When a table has bracket column names and stripBracketColumnNames=true,
    /// the columns should be renamed and BracketColumnsRenamed populated.
    /// </summary>
    [Fact]
    [Trait("Layer", "Service")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Table")]
    [Trait("Feature", "DataModel")]
    [Trait("RequiresExcel", "true")]
    [Trait("Speed", "Medium")]
    public void AddToDataModel_BracketColumns_WithStrip_RenamesColumnsAndReturnsRenamed()
    {
        // Arrange
        CreateTableWithBracketColumns();

        // Act
        var batch = _fixture.BatchToken;
        var result = AddToDataModel("BracketTable", stripBracketColumnNames: true);

        // Assert
        Assert.True(result.Success, $"AddToDataModel failed: {result.ErrorMessage}");
        Assert.NotNull(result.BracketColumnsRenamed);
        Assert.Equal(2, result.BracketColumnsRenamed.Length);
        Assert.Contains("[ACR_CM1]", result.BracketColumnsRenamed);
        Assert.Contains("[ACR_CM2]", result.BracketColumnsRenamed);
        Assert.Null(result.BracketColumnsFound);

        // Verify the source column headers were actually renamed (brackets removed)
        var rangeResult = RequireSuccess(_commands.GetValues(batch, "Data", "A1:C1"));
        Assert.NotNull(rangeResult);
        Assert.Equal("ProductName", rangeResult.Values[0][0]?.ToString());
        Assert.Equal("ACR_CM1", rangeResult.Values[0][1]?.ToString());
        Assert.Equal("ACR_CM2", rangeResult.Values[0][2]?.ToString());
        AssertLoadedTable("BracketTable", ["ProductName", "ACR_CM1", "ACR_CM2"],
            [["Widget", 100, 200], ["Gadget", 150, 250]]);
    }

    /// <summary>
    /// When a table has no bracket column names, BracketColumnsFound and BracketColumnsRenamed
    /// should both be null.
    /// </summary>
    [Fact]
    [Trait("Layer", "Service")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Table")]
    [Trait("Feature", "DataModel")]
    [Trait("RequiresExcel", "true")]
    [Trait("Speed", "Medium")]
    public void AddToDataModel_NoBracketColumns_NoBracketFields()
    {
        // Arrange
        CreateTableWithNormalColumns();

        // Act
        var result = AddToDataModel("NormalTable", stripBracketColumnNames: false);

        // Assert
        Assert.True(result.Success, $"AddToDataModel failed: {result.ErrorMessage}");
        Assert.Null(result.BracketColumnsFound);
        Assert.Null(result.BracketColumnsRenamed);
        AssertLoadedTable("NormalTable", ["ProductName", "Amount"], [["Widget", 100]]);
    }

    /// <summary>
    /// Adding the same table to the Data Model twice should throw InvalidOperationException.
    /// </summary>
    [Fact]
    [Trait("Layer", "Service")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Table")]
    [Trait("Feature", "DataModel")]
    [Trait("RequiresExcel", "true")]
    [Trait("Speed", "Medium")]
    public void AddToDataModel_AlreadyInModel_ThrowsInvalidOperationException()
    {
        // Arrange
        CreateTableWithNormalColumns();

        var batch = _fixture.BatchToken;

        // First add succeeds
        var first = AddToDataModel("NormalTable");
        Assert.True(first.Success, $"First AddToDataModel failed: {first.ErrorMessage}");

        // Second add should throw (table already in model)
        var before = System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_model.ReadTable(batch, "NormalTable")));
        var error = Assert.Throws<InvalidOperationException>(() => AddToDataModel("NormalTable"));
        Assert.Contains("already in the Data Model", error.Message, StringComparison.Ordinal);
        Assert.Equal(before, System.Text.Json.JsonSerializer.Serialize(
            RequireSuccess(_model.ReadTable(batch, "NormalTable"))));
        AssertLoadedTable("NormalTable", ["ProductName", "Amount"], [["Widget", 100]]);
    }

    private void AssertLoadedTable(string name, string[] headers, object[][] rows)
    {
        var table = RequireSuccess(_model.ReadTable(_fixture.BatchToken, name));
        Assert.Equal(name, table.TableName);
        Assert.Equal(rows.Length, table.RecordCount);
        Assert.Equal(headers.Order(StringComparer.Ordinal),
            table.Columns.Select(column => column.Name).Order(StringComparer.Ordinal));
        var source = RequireSuccess(_tableCommands.GetData(_fixture.BatchToken, name));
        Assert.Equal(headers, source.Headers);
        Assert.Equal(rows.Length, source.Data.Count);
        for (var index = 0; index < rows.Length; index++)
            Assert.Equal(rows[index], source.Data[index]);

        var modelRows = RequireSuccess(_model.Evaluate(_fixture.BatchToken,
            $"EVALUATE '{name}' ORDER BY '{name}'[ProductName]"));
        Assert.Equal(rows.Length, modelRows.RowCount);
        Assert.Equal(headers.Length, modelRows.ColumnCount);
        var sorted = rows.OrderBy(row => (string)row[0], StringComparer.Ordinal).ToArray();
        for (var row = 0; row < sorted.Length; row++)
        {
            for (var column = 0; column < headers.Length; column++)
            {
                var modelColumn = modelRows.Columns.FindIndex(value => value == $"{name}[{headers[column]}]");
                Assert.InRange(modelColumn, 0, headers.Length - 1);
                var value = modelRows.Rows[row][modelColumn];
                if (column == 0)
                    Assert.Equal(sorted[row][column], value);
                else
                    Assert.Equal(Convert.ToDecimal(sorted[row][column], System.Globalization.CultureInfo.InvariantCulture),
                        Convert.ToDecimal(value, System.Globalization.CultureInfo.InvariantCulture));
            }
        }
    }
}
