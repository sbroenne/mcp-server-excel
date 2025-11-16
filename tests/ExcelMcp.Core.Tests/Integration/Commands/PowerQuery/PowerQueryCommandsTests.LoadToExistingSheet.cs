using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Regression tests for Issue #170: LoadTo silently fails when worksheet with target name already exists
/// Tests verify that LoadToAsync detects existing sheets and returns clear error message requiring explicit deletion
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public partial class PowerQueryCommandsTests
{
    #region LoadToAsync with Existing Sheet Tests (Issue #170)

    [Fact]
    public async Task LoadTo_LoadToTableWithExistingSheet_ReturnsErrorRequiringExplicitDeletion()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_LoadToTableWithExistingSheet_ReturnsErrorRequiringExplicitDeletion),
            _tempDir);
        var queryName = "TestLoadToExisting";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_LoadToTableWithExistingSheet_ReturnsErrorRequiringExplicitDeletion));
        var targetSheet = "ExistingSheet";

        // Act - Create query, sheet, then attempt LoadTo (all in one batch)
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection-only query
        await _powerQueryCommands.Create(
            batch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        // Create sheet that will conflict
        await _sheetCommands.Create(batch, targetSheet);

        // LoadTo should detect existing sheet and return error
        var loadResult = _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);

        // Assert - Should fail with clear error message
        Assert.False(loadResult.Success, "LoadTo should fail when sheet already exists");
        Assert.NotNull(loadResult.ErrorMessage);
        Assert.Contains("worksheet already exists", loadResult.ErrorMessage);
    }

    [Fact]
    public async Task LoadTo_LoadToBothWithExistingSheet_ReturnsErrorRequiringExplicitDeletion()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_LoadToBothWithExistingSheet_ReturnsErrorRequiringExplicitDeletion),
            _tempDir);
        var queryName = "TestLoadToBothExisting";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_LoadToBothWithExistingSheet_ReturnsErrorRequiringExplicitDeletion));
        var targetSheet = "ExistingSheetBoth";

        // Act - Create query, sheet, then attempt LoadTo (all in one batch)
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection-only query
        await _powerQueryCommands.Create(
            batch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        // Create sheet that will conflict
        await _sheetCommands.Create(batch, targetSheet);

        // LoadTo with LoadToBoth should detect existing sheet
        var loadResult = _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToBoth, targetSheet);

        // Assert - Should fail with clear error message
        Assert.False(loadResult.Success, "LoadToBoth should fail when sheet already exists");
        Assert.NotNull(loadResult.ErrorMessage);
        Assert.Contains("worksheet already exists", loadResult.ErrorMessage);
    }

    [Fact]
    public async Task LoadTo_DeleteExistingSheetThenLoadTo_SucceedsAfterManualDeletion()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_DeleteExistingSheetThenLoadTo_SucceedsAfterManualDeletion),
            _tempDir);
        var queryName = "TestSequentialLoadTo";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_DeleteExistingSheetThenLoadTo_SucceedsAfterManualDeletion));
        var targetSheet = "SequentialSheet";

        // Act - All in one batch: Create query, LoadTo, verify error, delete sheet, LoadTo succeeds
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection-only query
        await _powerQueryCommands.Create(
            batch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        // First LoadTo creates sheet with data
        var load1 = await _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);
        Assert.True(load1.Success, $"First LoadTo failed: {load1.ErrorMessage}");

        // Second LoadTo should fail because sheet exists
        var load2 = await _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);
        Assert.False(load2.Success, "Second LoadTo should fail when sheet already exists");
        Assert.Contains("worksheet already exists", load2.ErrorMessage);

        // User deletes sheet manually
        await _sheetCommands.Delete(batch, targetSheet);

        // LoadTo after manual deletion succeeds
        var load3 = await _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);
        Assert.True(load3.Success, $"LoadTo after manual deletion failed: {load3.ErrorMessage}");
        Assert.True(load3.ConfigurationApplied);
        Assert.True(load3.DataRefreshed);
        Assert.True(load3.RowsLoaded > 0);
    }

    [Fact]
    public async Task LoadTo_WithExistingSheetSameName_ReturnsErrorRequiringExplicitDeletion()
    {
        // Arrange - This test verifies the actual bug scenario: query and sheet have same name
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_WithExistingSheetSameName_ReturnsErrorRequiringExplicitDeletion),
            _tempDir);
        var queryName = "MilestoneExport"; // Use same name for query and sheet (bug scenario)
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_WithExistingSheetSameName_ReturnsErrorRequiringExplicitDeletion));

        // Act - Create sheet, query, then attempt LoadTo (all in one batch)
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create sheet with same name as query first
        await _sheetCommands.Create(batch, queryName);

        // Create connection-only query with same name
        await _powerQueryCommands.Create(
            batch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        // LoadTo with sheet already existing (bug report scenario)
        var loadResult = _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, queryName);

        // Assert - The bug was silent failure. Now we return clear error.
        Assert.False(loadResult.Success, "LoadTo should fail when target sheet already exists");
        Assert.NotNull(loadResult.ErrorMessage);
        Assert.Contains("worksheet already exists", loadResult.ErrorMessage);
    }

    [Fact]
    public async Task LoadTo_AfterManualDeletion_DataLoadedSuccessfully()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_AfterManualDeletion_DataLoadedSuccessfully),
            _tempDir);
        var queryName = "RoundTripQuery";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_AfterManualDeletion_DataLoadedSuccessfully));
        var targetSheet = "RoundTripSheet";

        // Act - All in one batch: Create query, sheet, verify error, delete, LoadTo succeeds
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create query (connection-only) and existing sheet
        await _powerQueryCommands.Create(
            batch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);
        await _sheetCommands.Create(batch, targetSheet);

        // Verify LoadTo fails with existing sheet
        var failResult = _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);
        Assert.False(failResult.Success, "LoadTo should fail with existing sheet");

        // Delete sheet manually
        await _sheetCommands.Delete(batch, targetSheet);

        // LoadTo after deletion succeeds
        var loadResult = _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);
        Assert.True(loadResult.Success, $"LoadTo after deletion failed: {loadResult.ErrorMessage}");
        Assert.True(loadResult.RowsLoaded > 0, "Data should be loaded");

        // Verify sheet exists with data
        var listSheets = await _sheetCommands.List(batch);
        Assert.Contains(listSheets.Worksheets, s => s.Name == targetSheet);
    }

    [Fact]
    public async Task LoadTo_NewSheetName_CreatesSheetSuccessfully()
    {
        // Arrange - Verify backwards compatibility: new sheet name still works
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(LoadTo_NewSheetName_CreatesSheetSuccessfully),
            _tempDir);
        var queryName = "TestNewSheet";
        var mCodeFile = CreateUniqueTestQueryFile(nameof(LoadTo_NewSheetName_CreatesSheetSuccessfully));
        var targetSheet = "BrandNewSheet";

        // Act - Create query and LoadTo with non-existent sheet (all in one batch)
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create connection-only query (no sheet exists)
        await _powerQueryCommands.Create(
            batch, queryName, mCodeFile, PowerQueryLoadMode.ConnectionOnly);

        // LoadTo with non-existent sheet should succeed
        var loadResult = _powerQueryCommands.LoadTo(
            batch, queryName, PowerQueryLoadMode.LoadToTable, targetSheet);

        // Assert
        Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");

        // Verify sheet was created
        var listResult = _sheetCommands.List(batch);
        Assert.Contains(listResult.Worksheets, s => s.Name == targetSheet);
    }

    #endregion
}
