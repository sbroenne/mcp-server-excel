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
    [Fact]
    public async Task Get_ExistingQueryTable_ReturnsDetails()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(Get_ExistingQueryTable_ReturnsDetails), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create QueryTable
        var dataModelCommands = new Sbroenne.ExcelMcp.Core.Commands.DataModelCommands();
        var pqCommands = new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Column1\"}, {{\"Data1\"}}) in Source";
        await pqCommands.ImportAsync(batch, "MyQuery", mCode);
        
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
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

    [Fact]
    public async Task Refresh_ExistingQueryTable_RefreshesSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(Refresh_ExistingQueryTable_RefreshesSuccessfully), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create QueryTable
        var dataModelCommands = new Sbroenne.ExcelMcp.Core.Commands.DataModelCommands();
        var pqCommands = new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"Name\"}, {{\"Test\"}}) in Source";
        await pqCommands.ImportAsync(batch, "RefreshQuery", mCode);
        
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "RefreshSheet");
        
        await _commands.CreateFromQueryAsync(batch, "RefreshSheet", "RefreshQT", "RefreshQuery");

        // Act
        var result = await _commands.RefreshAsync(batch, "RefreshQT");

        // Assert
        Assert.True(result.Success, $"Refresh failed: {result.ErrorMessage}");
    }

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

    [Fact]
    public async Task RefreshAll_MultipleQueryTables_RefreshesAll()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(RefreshAll_MultipleQueryTables_RefreshesAll), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create multiple QueryTables
        var dataModelCommands = new Sbroenne.ExcelMcp.Core.Commands.DataModelCommands();
        var pqCommands = new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode1 = "let Source = #table({\"A\"}, {{1}}) in Source";
        var mCode2 = "let Source = #table({\"B\"}, {{2}}) in Source";
        await pqCommands.ImportAsync(batch, "Q1", mCode1);
        await pqCommands.ImportAsync(batch, "Q2", mCode2);
        
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
        await sheetCommands.CreateAsync(batch, "S1");
        await sheetCommands.CreateAsync(batch, "S2");
        
        await _commands.CreateFromQueryAsync(batch, "S1", "QT1", "Q1");
        await _commands.CreateFromQueryAsync(batch, "S2", "QT2", "Q2");

        // Act
        var result = await _commands.RefreshAllAsync(batch);

        // Assert
        Assert.True(result.Success, $"RefreshAll failed: {result.ErrorMessage}");
    }

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

    [Fact]
    public async Task UpdateProperties_ExistingQueryTable_UpdatesSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(QueryTableCommandsTests), nameof(UpdateProperties_ExistingQueryTable_UpdatesSuccessfully), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create QueryTable with default options
        var dataModelCommands = new Sbroenne.ExcelMcp.Core.Commands.DataModelCommands();
        var pqCommands = new Sbroenne.ExcelMcp.Core.Commands.PowerQueryCommands(dataModelCommands);
        var mCode = "let Source = #table({\"X\"}, {{\"Y\"}}) in Source";
        await pqCommands.ImportAsync(batch, "UpdateQuery", mCode);
        
        var sheetCommands = new Sbroenne.ExcelMcp.Core.Commands.SheetCommands();
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
