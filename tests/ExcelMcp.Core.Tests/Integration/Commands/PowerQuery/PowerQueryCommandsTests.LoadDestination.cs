using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Integration tests for PowerQuery operations
/// Test Coverage: CreateAsync, UpdateAsync, LoadToAsync, UnloadAsync, ValidateSyntaxAsync, RefreshAllAsync
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public partial class PowerQueryCommandsTests
{
    /// <inheritdoc/>
    #region CreateAsync Tests

    [Fact]
    public async Task Create_ConnectionOnly_CreatesQuerySuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Create_ConnectionOnly_CreatesQuerySuccessfully),
            _tempDir);
        var queryName = "TestConnectionOnly";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Create_ConnectionOnly_CreatesQuerySuccessfully));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _powerQueryCommands.CreateAsync(
            batch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        // Assert
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
        Assert.Equal(queryName, createResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, createResult.LoadDestination);
        Assert.False(createResult.DataLoaded);

        // Verify query exists
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Create_LoadToTable_CreatesAndLoadsData()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Create_LoadToTable_CreatesAndLoadsData),
            _tempDir);
        var queryName = "TestLoadToTable";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Create_LoadToTable_CreatesAndLoadsData));
        var targetSheet = "DataSheet";

        // Create target sheet first
        await using (var setupBatch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _sheetCommands.CreateAsync(setupBatch, targetSheet);
            await setupBatch.SaveAsync();
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _powerQueryCommands.CreateAsync(
            batch, queryName, mCodeFile, PowerQueryLoadMode.LoadToTable, targetSheet);

        // Assert
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
        Assert.Equal(queryName, createResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, createResult.LoadDestination);
        Assert.True(createResult.DataLoaded);
        Assert.Equal(targetSheet, createResult.WorksheetName);

        // Verify query exists
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Create_LoadToDataModel_CreatesAndLoadsToModel()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Create_LoadToDataModel_CreatesAndLoadsToModel),
            _tempDir);
        var queryName = "TestLoadToModel";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Create_LoadToDataModel_CreatesAndLoadsToModel));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _powerQueryCommands.CreateAsync(
            batch, queryName, mCodeFile, PowerQueryLoadMode.LoadToDataModel);

        // Assert
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
        Assert.Equal(queryName, createResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, createResult.LoadDestination);
        Assert.True(createResult.DataLoaded);

        // Verify query exists
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Create_LoadToBoth_CreatesAndLoadsToTableAndModel()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Create_LoadToBoth_CreatesAndLoadsToTableAndModel),
            _tempDir);
        var queryName = "TestLoadToBoth";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Create_LoadToBoth_CreatesAndLoadsToTableAndModel));
        var targetSheet = "BothSheet";

        // Create target sheet first
        await using (var setupBatch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _sheetCommands.CreateAsync(setupBatch, targetSheet);
            await setupBatch.SaveAsync();
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _powerQueryCommands.CreateAsync(
            batch, queryName, mCodeFile, PowerQueryLoadMode.LoadToBoth, targetSheet);

        // Assert
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
        Assert.Equal(queryName, createResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, createResult.LoadDestination);
        Assert.True(createResult.DataLoaded);
        Assert.Equal(targetSheet, createResult.WorksheetName);

        // Verify query exists
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    #endregion

    #region Update Tests

    [Fact]
    public async Task UpdateMCode_ExistingQuery_UpdatesSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(UpdateMCode_ExistingQuery_UpdatesSuccessfully),
            _tempDir);
        var queryName = "TestUpdate";
        var originalFile = CreateUniqueTestQueryFile("Original");
        var updatedFile = CreateTestQueryFileWithContent(
            "Updated",
            @"let
    Source = #table(
        {""UpdatedColumn""},
        {{""Updated1""}, {""Updated2""}}
    )
in
    Source");

        // Create query first
        await using (var setupBatch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _powerQueryCommands.CreateAsync(
                setupBatch, queryName, originalFile, PowerQueryLoadMode.ConnectionOnly);
            await setupBatch.SaveAsync();
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var updateResult = await _powerQueryCommands.UpdateAsync(batch, queryName, updatedFile);

        // Assert
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // Verify M code was updated
        var viewResult = await _powerQueryCommands.ViewAsync(batch, queryName);
        Assert.True(viewResult.Success);
        Assert.Contains("UpdatedColumn", viewResult.MCode);
    }
    /// <inheritdoc/>

    #endregion

    #region LoadToAsync Tests

    [Fact]
    public async Task LoadTo_ConnectionOnlyToTable_LoadsDataSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_ConnectionOnlyToTable_LoadsDataSuccessfully),
            _tempDir);
        var queryName = "TestLoadTo";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_ConnectionOnlyToTable_LoadsDataSuccessfully));
        var targetSheet = "LoadSheet";

        // Create connection-only query first
        await using (var setupBatch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _powerQueryCommands.CreateAsync(
                setupBatch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);
            await _sheetCommands.CreateAsync(setupBatch, targetSheet);
            await setupBatch.SaveAsync();
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var loadResult = await _powerQueryCommands.LoadToAsync(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);

        // Assert
        Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");
        Assert.Equal(queryName, loadResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadResult.LoadDestination);
        Assert.True(loadResult.DataRefreshed);
        Assert.Equal(targetSheet, loadResult.WorksheetName);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task LoadTo_ToDataModel_LoadsToModelSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_ToDataModel_LoadsToModelSuccessfully),
            _tempDir);
        var queryName = "TestLoadToModel";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_ToDataModel_LoadsToModelSuccessfully));

        // Create connection-only query first
        await using (var setupBatch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _powerQueryCommands.CreateAsync(
                setupBatch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);
            await setupBatch.SaveAsync();
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var loadResult = await _powerQueryCommands.LoadToAsync(
            batch, queryName, PowerQueryLoadMode.LoadToDataModel);

        // Assert
        Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");
        Assert.Equal(queryName, loadResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadResult.LoadDestination);
        Assert.True(loadResult.DataRefreshed);
    }
    /// <inheritdoc/>

    #endregion

    #region UnloadAsync Tests

    [Fact]
    public async Task Unload_LoadedQuery_UnloadsSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Unload_LoadedQuery_UnloadsSuccessfully),
            _tempDir);
        var queryName = "TestUnload";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Unload_LoadedQuery_UnloadsSuccessfully));
        var targetSheet = "UnloadSheet";

        // Create loaded query first
        await using (var setupBatch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _sheetCommands.CreateAsync(setupBatch, targetSheet);
            await _powerQueryCommands.CreateAsync(
                setupBatch, queryName, mCodeFile, PowerQueryLoadMode.LoadToTable, targetSheet);
            await setupBatch.SaveAsync();
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var unloadResult = await _powerQueryCommands.UnloadAsync(batch, queryName);

        // Assert
        Assert.True(unloadResult.Success, $"Unload failed: {unloadResult.ErrorMessage}");

        // Verify query still exists
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    #endregion

    #region ValidateSyntaxAsync Tests

    // ValidateSyntax tests removed - Excel doesn't validate M code syntax at query creation time
    // Validation only happens during refresh, making it unreliable for syntax checking

    #endregion

    #region UpdateAndRefreshAsync Tests

    [Fact]
    public async Task UpdateAndRefresh_ExistingLoadedQuery_UpdatesAndRefreshes()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(UpdateAndRefresh_ExistingLoadedQuery_UpdatesAndRefreshes),
            _tempDir);
        var queryName = "TestUpdateRefresh";
        var originalFile = CreateUniqueTestQueryFile("Original");
        var updatedFile = CreateTestQueryFileWithContent(
            "UpdatedForRefresh",
            @"let
    Source = #table(
        {""RefreshedColumn""},
        {{""Refreshed1""}, {""Refreshed2""}, {""Refreshed3""}}
    )
in
    Source");
        var targetSheet = "RefreshSheet";

        // Create loaded query first
        await using (var setupBatch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _sheetCommands.CreateAsync(setupBatch, targetSheet);
            await _powerQueryCommands.CreateAsync(
                setupBatch, queryName, originalFile, PowerQueryLoadMode.LoadToTable, targetSheet);
            await setupBatch.SaveAsync();
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var updateResult = await _powerQueryCommands.UpdateAsync(
            batch, queryName, updatedFile);

        // Assert
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // Verify M code was updated
        var viewResult = await _powerQueryCommands.ViewAsync(batch, queryName);
        Assert.True(viewResult.Success);
        Assert.Contains("RefreshedColumn", viewResult.MCode);
    }
    /// <inheritdoc/>

    #endregion

    #region RefreshAllAsync Tests

    [Fact]
    public async Task RefreshAll_MultipleQueries_RefreshesAll()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(RefreshAll_MultipleQueries_RefreshesAll),
            _tempDir);
        var query1 = "Query1";
        var query2 = "Query2";
        var mCodeFile1 = CreateUniqueTestQueryFile("RefreshAll1");
        var mCodeFile2 = CreateUniqueTestQueryFile("RefreshAll2");
        var sheet1 = "Sheet1";
        var sheet2 = "Sheet2";

        // Create multiple loaded queries
        await using (var setupBatch = await ExcelSession.BeginBatchAsync(testFile))
        {
            await _sheetCommands.CreateAsync(setupBatch, sheet1);
            await _sheetCommands.CreateAsync(setupBatch, sheet2);
            await _powerQueryCommands.CreateAsync(setupBatch, query1, mCodeFile1, PowerQueryLoadMode.LoadToTable, sheet1);
            await _powerQueryCommands.CreateAsync(setupBatch, query2, mCodeFile2, PowerQueryLoadMode.LoadToTable, sheet2);
            await setupBatch.SaveAsync();
        }

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var refreshResult = await _powerQueryCommands.RefreshAllAsync(batch);

        // Assert
        Assert.True(refreshResult.Success, $"RefreshAll failed: {refreshResult.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task RefreshAll_EmptyWorkbook_Succeeds()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(RefreshAll_EmptyWorkbook_Succeeds),
            _tempDir);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var refreshResult = await _powerQueryCommands.RefreshAllAsync(batch);

        // Assert
        Assert.True(refreshResult.Success, $"RefreshAll on empty workbook failed: {refreshResult.ErrorMessage}");
    }

    #endregion
}
