using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Integration tests for TableCommands.AddToDataModel, focusing on bracket column name detection
/// and stripping. Regression tests for the stripBracketColumnNames feature.
/// </summary>
public class TableCommandsDataModelTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TableCommands _tableCommands;
    private readonly RangeCommands _rangeCommands;
    private readonly TempDirectoryFixture _fixture;

    public TableCommandsDataModelTests(TempDirectoryFixture fixture)
    {
        _tableCommands = new TableCommands();
        _rangeCommands = new RangeCommands();
        _fixture = fixture;
    }

    private static string CreateTableWithBracketColumns(string filePath)
    {
        using var batch = ExcelSession.BeginBatch(filePath);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Name = "Data";
            // Column A: normal name, Column B/C: bracket names
            sheet.Range["A1"].Value2 = "ProductName";
            sheet.Range["B1"].Value2 = "[ACR_CM1]";
            sheet.Range["C1"].Value2 = "[ACR_CM2]";
            sheet.Range["A2"].Value2 = "Widget";
            sheet.Range["B2"].Value2 = 100.0;
            sheet.Range["C2"].Value2 = 200.0;
            sheet.Range["A3"].Value2 = "Gadget";
            sheet.Range["B3"].Value2 = 150.0;
            sheet.Range["C3"].Value2 = 250.0;
        });
        var tableResult = new TableCommands().Create(batch, "Data", "BracketTable", "A1:C3");
        Assert.True(tableResult.Success, $"Setup failed: {tableResult.ErrorMessage}");
        batch.Save();
        return "BracketTable";
    }

    private static string CreateTableWithNormalColumns(string filePath)
    {
        using var batch = ExcelSession.BeginBatch(filePath);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Name = "Data";
            sheet.Range["A1"].Value2 = "ProductName";
            sheet.Range["B1"].Value2 = "Amount";
            sheet.Range["A2"].Value2 = "Widget";
            sheet.Range["B2"].Value2 = 100.0;
        });
        var tableResult = new TableCommands().Create(batch, "Data", "NormalTable", "A1:B2");
        Assert.True(tableResult.Success, $"Setup failed: {tableResult.ErrorMessage}");
        batch.Save();
        return "NormalTable";
    }

    /// <summary>
    /// When a table has bracket column names and stripBracketColumnNames=false,
    /// BracketColumnsFound should be populated with the bracket column names.
    /// </summary>
    [Fact]
    [Trait("Layer", "Core")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Table")]
    [Trait("Feature", "DataModel")]
    [Trait("RequiresExcel", "true")]
    [Trait("Speed", "Medium")]
    public void AddToDataModel_BracketColumns_WithoutStrip_ReturnsBracketColumnsFound()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        CreateTableWithBracketColumns(testFile);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _tableCommands.AddToDataModel(batch, "BracketTable", stripBracketColumnNames: false);

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
    [Trait("Layer", "Core")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Table")]
    [Trait("Feature", "DataModel")]
    [Trait("RequiresExcel", "true")]
    [Trait("Speed", "Medium")]
    public void AddToDataModel_BracketColumns_WithStrip_RenamesColumnsAndReturnsRenamed()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        CreateTableWithBracketColumns(testFile);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _tableCommands.AddToDataModel(batch, "BracketTable", stripBracketColumnNames: true);

        // Assert
        Assert.True(result.Success, $"AddToDataModel failed: {result.ErrorMessage}");
        Assert.NotNull(result.BracketColumnsRenamed);
        Assert.Equal(2, result.BracketColumnsRenamed.Length);
        Assert.Contains("[ACR_CM1]", result.BracketColumnsRenamed);
        Assert.Contains("[ACR_CM2]", result.BracketColumnsRenamed);
        Assert.Null(result.BracketColumnsFound);

        // Verify the source column headers were actually renamed (brackets removed)
        var rangeResult = _rangeCommands.GetValues(batch, "Data", "A1:C1");
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
    [Trait("Layer", "Core")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Table")]
    [Trait("Feature", "DataModel")]
    [Trait("RequiresExcel", "true")]
    [Trait("Speed", "Medium")]
    public void AddToDataModel_NoBracketColumns_NoBracketFields()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        CreateTableWithNormalColumns(testFile);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _tableCommands.AddToDataModel(batch, "NormalTable", stripBracketColumnNames: false);

        // Assert
        Assert.True(result.Success, $"AddToDataModel failed: {result.ErrorMessage}");
        Assert.Null(result.BracketColumnsFound);
        Assert.Null(result.BracketColumnsRenamed);
    }

    /// <summary>
    /// Adding the same table to the Data Model twice should throw InvalidOperationException.
    /// </summary>
    [Fact]
    [Trait("Layer", "Core")]
    [Trait("Category", "Integration")]
    [Trait("Feature", "Table")]
    [Trait("Feature", "DataModel")]
    [Trait("RequiresExcel", "true")]
    [Trait("Speed", "Medium")]
    public void AddToDataModel_AlreadyInModel_ThrowsInvalidOperationException()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();
        CreateTableWithNormalColumns(testFile);

        using var batch = ExcelSession.BeginBatch(testFile);

        // First add succeeds
        var first = _tableCommands.AddToDataModel(batch, "NormalTable");
        Assert.True(first.Success, $"First AddToDataModel failed: {first.ErrorMessage}");

        // Second add should throw (table already in model)
        Assert.ThrowsAny<Exception>(() => _tableCommands.AddToDataModel(batch, "NormalTable"));
    }
}
