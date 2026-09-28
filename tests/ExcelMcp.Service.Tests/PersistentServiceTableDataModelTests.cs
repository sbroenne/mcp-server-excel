using Sbroenne.ExcelMcp.Core.Commands.Table;
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

    private void CreateTableWithBracketColumns()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        _commands.SetValues(
            batch,
            "Data",
            "A1:C3",
            [
                ["ProductName", "[ACR_CM1]", "[ACR_CM2]"],
                ["Widget", 100.0, 200.0],
                ["Gadget", 150.0, 250.0]
            ]);
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
        _commands.SetValues(
            batch,
            "Data",
            "A1:B2",
            [["ProductName", "Amount"], ["Widget", 100.0]]);
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
        var rangeResult = _commands.GetValues(batch, "Data", "A1:C1");
        Assert.NotNull(rangeResult);
        Assert.Equal("ProductName", rangeResult.Values[0][0]?.ToString());
        Assert.Equal("ACR_CM1", rangeResult.Values[0][1]?.ToString());
        Assert.Equal("ACR_CM2", rangeResult.Values[0][2]?.ToString());
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
        Assert.ThrowsAny<Exception>(() => AddToDataModel("NormalTable"));
    }
}
