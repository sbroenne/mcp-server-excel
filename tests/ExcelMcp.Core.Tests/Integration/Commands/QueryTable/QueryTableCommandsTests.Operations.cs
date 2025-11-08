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
    public async Task Get_ExistingQueryTable_ReturnsDetails()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(Get_ExistingQueryTable_ReturnsDetails), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create QueryTable
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Column1\"}, {{\"Data1\"}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "MyQuery.pq");
        await System.IO.File.WriteAllTextAsync(mCodeFile, mCode);
        await pqCommands.CreateAsync(batch, "MyQuery", mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "Sheet1");

        await _commands.CreateFromQueryAsync(batch, "Sheet1", "MyQT", "MyQuery");

        // Act
        var result = await _commands.GetAsync(batch, "MyQT");

        // Assert
        Assert.True(result.Success, $"Get failed: {result.ErrorMessage}");
        Assert.NotNull(result.QueryTable);
        Assert.Equal("MyQT", result.QueryTable.Name);
        Assert.Equal("Sheet1", result.QueryTable.WorksheetName);
        Assert.NotEmpty(result.QueryTable.Range);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Get_NonExistentQueryTable_ReturnsFalse()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(Get_NonExistentQueryTable_ReturnsFalse), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _commands.GetAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Refresh_ExistingQueryTable_RefreshesSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(Refresh_ExistingQueryTable_RefreshesSuccessfully), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create QueryTable
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Name\"}, {{\"Test\"}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "RefreshQuery.pq");
        await System.IO.File.WriteAllTextAsync(mCodeFile, mCode);
        await pqCommands.CreateAsync(batch, "RefreshQuery", mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "RefreshSheet");

        await _commands.CreateFromQueryAsync(batch, "RefreshSheet", "RefreshQT", "RefreshQuery");

        // Act
        var result = await _commands.RefreshAsync(batch, "RefreshQT");

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Refresh_NonExistentQueryTable_ReturnsFalse()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(Refresh_NonExistentQueryTable_ReturnsFalse), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _commands.RefreshAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task RefreshAll_MultipleQueryTables_RefreshesAll()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(RefreshAll_MultipleQueryTables_RefreshesAll), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create multiple QueryTables
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode1 = "let Source = #table({\"A\"}, {{1}}) in Source";
        var mCode2 = "let Source = #table({\"B\"}, {{2}}) in Source";
        var mCodeFile1 = Path.Combine(_tempDir, "Q1.pq");
        var mCodeFile2 = Path.Combine(_tempDir, "Q2.pq");
        await System.IO.File.WriteAllTextAsync(mCodeFile1, mCode1);
        await System.IO.File.WriteAllTextAsync(mCodeFile2, mCode2);
        await pqCommands.CreateAsync(batch, "Q1", mCodeFile1, PowerQueryLoadMode.ConnectionOnly);
        await pqCommands.CreateAsync(batch, "Q2", mCodeFile2, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "S1");
        await sheetCommands.CreateAsync(batch, "S2");

        await _commands.CreateFromQueryAsync(batch, "S1", "QT1", "Q1");
        await _commands.CreateFromQueryAsync(batch, "S2", "QT2", "Q2");

        // Act
        var result = await _commands.RefreshAllAsync(batch);

        // Assert
        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task RefreshAll_EmptyWorkbook_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(RefreshAll_EmptyWorkbook_ReturnsSuccess), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Act
        var result = await _commands.RefreshAllAsync(batch);

        // Assert
        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task UpdateProperties_ExistingQueryTable_UpdatesSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(UpdateProperties_ExistingQueryTable_UpdatesSuccessfully), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create QueryTable with default options
        var dataModelCommands = new Core.Commands.DataModelCommands();
        var pqCommands = new Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"X\"}, {{\"Y\"}}) in Source";
        var mCodeFile = Path.Combine(_tempDir, "UpdateQuery.pq");
        await System.IO.File.WriteAllTextAsync(mCodeFile, mCode);
        await pqCommands.CreateAsync(batch, "UpdateQuery", mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        var sheetCommands = new Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "UpdateSheet");

        await _commands.CreateFromQueryAsync(batch, "UpdateSheet", "UpdateQT", "UpdateQuery");

        // Act
        var updateOptions = new QueryTableUpdateOptions
        {
            BackgroundQuery = true,
            RefreshOnFileOpen = true
        };
        var result = await _commands.UpdatePropertiesAsync(batch, "UpdateQT", updateOptions);

        // Assert
        Assert.True(result.Success, $"UpdateProperties failed: {result.ErrorMessage}");

        // Verify updated properties
        var getResult = await _commands.GetAsync(batch, "UpdateQT");
        Assert.True(getResult.Success);
        Assert.NotNull(getResult.QueryTable);
        Assert.True(getResult.QueryTable.BackgroundQuery);
        Assert.True(getResult.QueryTable.RefreshOnFileOpen);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task UpdateProperties_NonExistentQueryTable_ReturnsFalse()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(UpdateProperties_NonExistentQueryTable_ReturnsFalse), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var updateOptions = new QueryTableUpdateOptions
        {
            BackgroundQuery = true
        };

        // Act
        var result = await _commands.UpdatePropertiesAsync(batch, "NonExistent", updateOptions);

        // Assert
        Assert.False(result.Success);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
}




