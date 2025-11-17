using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Integration tests for PowerQuery operations
/// Test Coverage: Create, Update, LoadToAsync, UnloadAsync, ValidateSyntaxAsync, RefreshAllAsync
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public partial class PowerQueryCommandsTests
{
    /// <inheritdoc/>
    #region Create Tests

    [Fact]
    public void Create_ConnectionOnly_CreatesQuerySuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Create_ConnectionOnly_CreatesQuerySuccessfully),
            _tempDir);
        var queryName = "TestConnectionOnly";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Create_ConnectionOnly_CreatesQuerySuccessfully));

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _powerQueryCommands.Create(
            batch, queryName, ReadMCodeFile(mCodeFile), PowerQueryLoadMode.ConnectionOnly);

        // Assert
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
        Assert.Equal(queryName, createResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, createResult.LoadDestination);
        Assert.False(createResult.DataLoaded);

        // Verify query exists
        var listResult = _powerQueryCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_LoadToTable_CreatesAndLoadsData()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Create_LoadToTable_CreatesAndLoadsData),
            _tempDir);
        var queryName = "TestLoadToTable";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Create_LoadToTable_CreatesAndLoadsData));
        var targetSheet = "DataSheet";

        // Create target sheet first
        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            _sheetCommands.Create(setupBatch, targetSheet);
            setupBatch.Save();
        }

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _powerQueryCommands.Create(
            batch, queryName, ReadMCodeFile(mCodeFile), PowerQueryLoadMode.LoadToTable, targetSheet);

        // Assert
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
        Assert.Equal(queryName, createResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, createResult.LoadDestination);
        Assert.True(createResult.DataLoaded);
        Assert.Equal(targetSheet, createResult.WorksheetName);

        // Verify query exists
        var listResult = _powerQueryCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_LoadToDataModel_CreatesAndLoadsToModel()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Create_LoadToDataModel_CreatesAndLoadsToModel),
            _tempDir);
        var queryName = "TestLoadToModel";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Create_LoadToDataModel_CreatesAndLoadsToModel));

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _powerQueryCommands.Create(
            batch, queryName, ReadMCodeFile(mCodeFile), PowerQueryLoadMode.LoadToDataModel);

        // Assert
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
        Assert.Equal(queryName, createResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, createResult.LoadDestination);
        Assert.True(createResult.DataLoaded);

        // Verify query exists
        var listResult = _powerQueryCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_LoadToBoth_CreatesAndLoadsToTableAndModel()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Create_LoadToBoth_CreatesAndLoadsToTableAndModel),
            _tempDir);
        var queryName = "TestLoadToBoth";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Create_LoadToBoth_CreatesAndLoadsToTableAndModel));
        var targetSheet = "BothSheet";

        // Create target sheet first
        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            _sheetCommands.Create(setupBatch, targetSheet);
            setupBatch.Save();
        }

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var createResult = _powerQueryCommands.Create(
            batch, queryName, ReadMCodeFile(mCodeFile), PowerQueryLoadMode.LoadToBoth, targetSheet);

        // Assert
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
        Assert.Equal(queryName, createResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, createResult.LoadDestination);
        Assert.True(createResult.DataLoaded);
        Assert.Equal(targetSheet, createResult.WorksheetName);

        // Verify query exists
        var listResult = _powerQueryCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    #endregion

    #region Update Tests

    [Fact]
    public void UpdateMCode_ExistingQuery_UpdatesSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
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
        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            _powerQueryCommands.Create(
                setupBatch, queryName, ReadMCodeFile(originalFile), PowerQueryLoadMode.ConnectionOnly);
            setupBatch.Save();
        }

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var updateResult = _powerQueryCommands.Update(batch, queryName, ReadMCodeFile(updatedFile));

        // Assert
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // Verify M code was updated
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success);
        Assert.Contains("UpdatedColumn", viewResult.MCode);
    }
    /// <inheritdoc/>

    #endregion

    #region LoadToAsync Tests

    [Fact]
    public void LoadTo_ConnectionOnlyToTable_LoadsDataSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_ConnectionOnlyToTable_LoadsDataSuccessfully),
            _tempDir);
        var queryName = "TestLoadTo";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_ConnectionOnlyToTable_LoadsDataSuccessfully));
        var targetSheet = "LoadSheet";

        // Act - Create connection-only query and LoadTo (all in one batch)
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection-only query first
        _powerQueryCommands.Create(
            batch, queryName, ReadMCodeFile(mCodeFile), PowerQueryLoadMode.ConnectionOnly);

        // LoadTo should create sheet and load data
        var loadResult = _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);

        // Assert
        Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");
        Assert.Equal(queryName, loadResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadResult.LoadDestination);
        Assert.True(loadResult.DataRefreshed);
        Assert.Equal(targetSheet, loadResult.WorksheetName);
        Assert.Equal("A1", loadResult.TargetCellAddress);
    }
    /// <inheritdoc/>

    [Fact]
    public void LoadTo_ToDataModel_LoadsToModelSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_ToDataModel_LoadsToModelSuccessfully),
            _tempDir);
        var queryName = "TestLoadToModel";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_ToDataModel_LoadsToModelSuccessfully));

        // Create connection-only query first
        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            _powerQueryCommands.Create(
                setupBatch, queryName, ReadMCodeFile(mCodeFile), PowerQueryLoadMode.ConnectionOnly);
            setupBatch.Save();
        }

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var loadResult = _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToDataModel);

        // Assert
        Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");
        Assert.Equal(queryName, loadResult.QueryName);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadResult.LoadDestination);
        Assert.True(loadResult.DataRefreshed);
        Assert.Null(loadResult.TargetCellAddress);
    }

    [Fact]
    public void LoadTo_ToDataModelWithTargetCell_ReturnsError()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_ToDataModelWithTargetCell_ReturnsError),
            _tempDir);
        var queryName = "TargetCellModel";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_ToDataModelWithTargetCell_ReturnsError));

        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            _powerQueryCommands.Create(
                setupBatch, queryName, ReadMCodeFile(mCodeFile), PowerQueryLoadMode.ConnectionOnly);
            setupBatch.Save();
        }

        using var batch = ExcelSession.BeginBatch(testFile);
        var loadResult = _powerQueryCommands.LoadTo(
            batch,
            queryName,
            PowerQueryLoadMode.LoadToDataModel,
            null,
            "B3");

        Assert.False(loadResult.Success, "LoadTo should reject targetCellAddress for Data Model loads");
        Assert.Contains("only supported", loadResult.ErrorMessage);
    }
    /// <inheritdoc/>

    #endregion

    #region UnloadAsync Tests

    [Fact]
    public void Unload_LoadedQuery_UnloadsSuccessfully()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Unload_LoadedQuery_UnloadsSuccessfully),
            _tempDir);
        var queryName = "TestUnload";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(Unload_LoadedQuery_UnloadsSuccessfully));
        var targetSheet = "UnloadSheet";

        // Create loaded query first
        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            _sheetCommands.Create(setupBatch, targetSheet);
            _powerQueryCommands.Create(
                setupBatch, queryName, ReadMCodeFile(mCodeFile), PowerQueryLoadMode.LoadToTable, targetSheet);
            setupBatch.Save();
        }

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var unloadResult = _powerQueryCommands.Unload(batch, queryName);

        // Assert
        Assert.True(unloadResult.Success, $"Unload failed: {unloadResult.ErrorMessage}");

        // Verify query still exists
        var listResult = _powerQueryCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Queries, q => q.Name == queryName);
    }
    /// <inheritdoc/>

    #endregion

    #region ValidateSyntaxAsync Tests

    // ValidateSyntax tests removed - Excel doesn't validate M code syntax at query creation time
    // Validation only happens during refresh, making it unreliable for syntax checking

    #endregion

    #region UpdateAndRefresh Tests

    [Fact]
    public void UpdateAndRefresh_ExistingLoadedQuery_UpdatesAndRefreshes()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
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
        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            _sheetCommands.Create(setupBatch, targetSheet);
            _powerQueryCommands.Create(
                setupBatch, queryName, ReadMCodeFile(originalFile), PowerQueryLoadMode.LoadToTable, targetSheet);
            setupBatch.Save();
        }

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var updateResult = _powerQueryCommands.Update(
            batch, queryName, ReadMCodeFile(updatedFile));

        // Assert
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // Verify M code was updated
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success);
        Assert.Contains("RefreshedColumn", viewResult.MCode);
    }
    /// <inheritdoc/>

    #endregion

    #region RefreshAllAsync Tests

    [Fact]
    public void RefreshAll_MultipleQueries_RefreshesAll()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
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
        using (var setupBatch = ExcelSession.BeginBatch(testFile))
        {
            _sheetCommands.Create(setupBatch, sheet1);
            _sheetCommands.Create(setupBatch, sheet2);
            _powerQueryCommands.Create(setupBatch, query1, ReadMCodeFile(mCodeFile1), PowerQueryLoadMode.LoadToTable, sheet1);
            _powerQueryCommands.Create(setupBatch, query2, ReadMCodeFile(mCodeFile2), PowerQueryLoadMode.LoadToTable, sheet2);
            setupBatch.Save();
        }

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var refreshResult = _powerQueryCommands.RefreshAll(batch);

        // Assert
        Assert.True(refreshResult.Success, $"RefreshAll failed: {refreshResult.ErrorMessage}");
    }
    /// <inheritdoc/>

    [Fact]
    public void RefreshAll_EmptyWorkbook_Succeeds()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(RefreshAll_EmptyWorkbook_Succeeds),
            _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var refreshResult = _powerQueryCommands.RefreshAll(batch);

        // Assert
        Assert.True(refreshResult.Success, $"RefreshAll on empty workbook failed: {refreshResult.ErrorMessage}");
    }

    #endregion
}
