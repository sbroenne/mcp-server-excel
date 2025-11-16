using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.QueryTable;

/// <summary>
/// Tests for QueryTable operations (Get, Refresh, UpdateProperties)
/// </summary>
public partial class QueryTableCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void Get_ExistingQueryTable_ReturnsDetails()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(Get_ExistingQueryTable_ReturnsDetails), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create QueryTable
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Column1\"}, {{\"Data1\"}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "MyQuery.pq");
        System.IO.File.WriteAllText(mCodeFile, mCode);
        pqCommands.Create(batch, "MyQuery", mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        sheetCommands.Create(batch, "Sheet1");

        _commands.CreateFromQuery(batch, "Sheet1", "MyQT", "MyQuery");

        // Act
        var result = _commands.Read(batch, "MyQT");

        // Assert
        Assert.True(result.Success, $"Get failed: {result.ErrorMessage}");
        Assert.NotNull(result.QueryTable);
        Assert.Equal("MyQT", result.QueryTable.Name);
        Assert.Equal("Sheet1", result.QueryTable.WorksheetName);
        Assert.NotEmpty(result.QueryTable.Range);
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_NonExistentQueryTable_ReturnsFalse()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(Get_NonExistentQueryTable_ReturnsFalse), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Read(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public void Refresh_ExistingQueryTable_RefreshesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(Refresh_ExistingQueryTable_RefreshesSuccessfully), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create QueryTable
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Name\"}, {{\"Test\"}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "RefreshQuery.pq");
        System.IO.File.WriteAllText(mCodeFile, mCode);
        pqCommands.Create(batch, "RefreshQuery", mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        sheetCommands.Create(batch, "RefreshSheet");

        _commands.CreateFromQuery(batch, "RefreshSheet", "RefreshQT", "RefreshQuery");

        // Act
        var result = _commands.Refresh(batch, "RefreshQT");

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void Refresh_NonExistentQueryTable_ReturnsFalse()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(Refresh_NonExistentQueryTable_ReturnsFalse), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.Refresh(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public void RefreshAll_MultipleQueryTables_RefreshesAll()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(RefreshAll_MultipleQueryTables_RefreshesAll), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create multiple QueryTables
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode1 = "let Source = #table({\"A\"}, {{1}}) in Source";
        var mCode2 = "let Source = #table({\"B\"}, {{2}}) in Source";
        var mCodeFile1 = Path.Combine(_tempDir, "Q1.pq");
        var mCodeFile2 = Path.Combine(_tempDir, "Q2.pq");
        System.IO.File.WriteAllText(mCodeFile1, mCode1);
        System.IO.File.WriteAllText(mCodeFile2, mCode2);
        pqCommands.Create(batch, "Q1", mCodeFile1, PowerQueryLoadMode.ConnectionOnly);
        pqCommands.Create(batch, "Q2", mCodeFile2, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        sheetCommands.Create(batch, "S1");
        sheetCommands.Create(batch, "S2");

        _commands.CreateFromQuery(batch, "S1", "QT1", "Q1");
        _commands.CreateFromQuery(batch, "S2", "QT2", "Q2");

        // Act
        var result = _commands.RefreshAll(batch);

        // Assert
        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void RefreshAll_EmptyWorkbook_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(RefreshAll_EmptyWorkbook_ReturnsSuccess), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var result = _commands.RefreshAll(batch);

        // Assert
        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void UpdateProperties_ExistingQueryTable_UpdatesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(UpdateProperties_ExistingQueryTable_UpdatesSuccessfully), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create QueryTable with default options
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"X\"}, {{\"Y\"}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "UpdateQuery.pq");
        System.IO.File.WriteAllText(mCodeFile, mCode);
        pqCommands.Create(batch, "UpdateQuery", mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        sheetCommands.Create(batch, "UpdateSheet");

        _commands.CreateFromQuery(batch, "UpdateSheet", "UpdateQT", "UpdateQuery");

        // Act
        var updateOptions = new QueryTableUpdateOptions
        {
            BackgroundQuery = true,
            RefreshOnFileOpen = true
        };
        var result = _commands.UpdateProperties(batch, "UpdateQT", updateOptions);

        // Assert
        Assert.True(result.Success, $"UpdateProperties failed: {result.ErrorMessage}");

        // Verify updated properties
        var getResult = _commands.Read(batch, "UpdateQT");
        Assert.True(getResult.Success);
        Assert.NotNull(getResult.QueryTable);
        Assert.True(getResult.QueryTable.BackgroundQuery);
        Assert.True(getResult.QueryTable.RefreshOnFileOpen);
    }
    /// <inheritdoc/>

    [Fact]
    public void UpdateProperties_NonExistentQueryTable_ReturnsFalse()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(QueryTableCommandsTests), nameof(UpdateProperties_NonExistentQueryTable_ReturnsFalse), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        var updateOptions = new QueryTableUpdateOptions
        {
            BackgroundQuery = true
        };

        // Act
        var result = _commands.UpdateProperties(batch, "NonExistent", updateOptions);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
}




