using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for Power Query operations with Data Model loading.
///
/// These tests validate that:
/// - LoadToDataModel settings are preserved after Update
/// - No duplicate tables are created in the Data Model
/// - Refresh operations work correctly with Data Model loading
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Feature", "PowerQuery")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public class PowerQueryDataModelLoadingTests : IClassFixture<TempDirectoryFixture>
{
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly DataModelCommands _dataModelCommands;
    private readonly TempDirectoryFixture _fixture;

    public PowerQueryDataModelLoadingTests(TempDirectoryFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(_dataModelCommands);
        _fixture = fixture;
    }

    #region LoadToDataModel Preservation Tests

    /// <summary>
    /// CRITICAL TEST: Validates that Update preserves LoadToDataModel settings
    /// and doesn't create duplicate tables in the Data Model.
    ///
    /// Scenario:
    /// 1. Create PowerQuery and load to Data Model
    /// 2. Verify one table exists in Data Model
    /// 3. Update the PowerQuery M code
    /// 4. Verify still only ONE table in Data Model (no duplicates)
    /// 5. Verify LoadToDataModel setting is preserved
    /// </summary>
    [Fact]
    public async Task Update_LoadedToDataModel_PreservesSettingsAndNoDuplicateTables()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_DataModel_" + Guid.NewGuid().ToString("N")[..8];

        var initialMCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Amount""},
        {
            {1, ""Alpha"", 100},
            {2, ""Beta"", 200}
        }
    )
in
    Source";

        var updatedMCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Amount""},
        {
            {1, ""Alpha"", 150},
            {2, ""Beta"", 250},
            {3, ""Gamma"", 350}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery and load to Data Model
        _powerQueryCommands.Create(batch, queryName, initialMCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Verify initial load configuration
        var loadConfigBefore = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigBefore.Success, $"GetLoadConfig before failed: {loadConfigBefore.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfigBefore.LoadMode);

        // STEP 3: Verify ONE table exists in Data Model
        var tablesBefore = await _dataModelCommands.ListTables(batch);
        Assert.True(tablesBefore.Success, $"ListTables before failed: {tablesBefore.ErrorMessage}");
        var queryTablesBefore = tablesBefore.Tables.Where(t => t.Name == queryName).ToList();
        Assert.Single(queryTablesBefore);

        // STEP 4: Update the M code (this triggers auto-refresh)
        _powerQueryCommands.Update(batch, queryName, updatedMCode);

        // STEP 5: Verify load configuration is PRESERVED after Update
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, $"GetLoadConfig after failed: {loadConfigAfter.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfigAfter.LoadMode);

        // STEP 6: Verify still ONLY ONE table in Data Model (no duplicates)
        var tablesAfter = await _dataModelCommands.ListTables(batch);
        Assert.True(tablesAfter.Success, $"ListTables after failed: {tablesAfter.ErrorMessage}");
        var queryTablesAfter = tablesAfter.Tables.Where(t => t.Name == queryName).ToList();
        Assert.Single(queryTablesAfter);

        // STEP 7: Verify M code was actually updated
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("Gamma", viewResult.MCode);
        Assert.Contains("350", viewResult.MCode);
    }

    /// <summary>
    /// Tests that multiple sequential updates don't create duplicate Data Model tables.
    /// </summary>
    [Fact]
    public async Task Update_MultipleUpdatesToDataModel_NoDuplicateTables()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_MultiUpdate_" + Guid.NewGuid().ToString("N")[..8];

        var mCodeV1 = @"let Source = #table({""Val""}, {{1}}) in Source";
        var mCodeV2 = @"let Source = #table({""Val""}, {{2}}) in Source";
        var mCodeV3 = @"let Source = #table({""Val""}, {{3}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create with LoadToDataModel
        _powerQueryCommands.Create(batch, queryName, mCodeV1, PowerQueryLoadMode.LoadToDataModel);

        // Update #1
        _powerQueryCommands.Update(batch, queryName, mCodeV2);

        // Update #2
        _powerQueryCommands.Update(batch, queryName, mCodeV3);

        // Verify still LoadToDataModel
        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfig.Success, $"GetLoadConfig failed: {loadConfig.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfig.LoadMode);

        // Verify only ONE table in Data Model
        var tables = await _dataModelCommands.ListTables(batch);
        Assert.True(tables.Success, $"ListTables failed: {tables.ErrorMessage}");
        var queryTables = tables.Tables.Where(t => t.Name == queryName).ToList();
        Assert.Single(queryTables);
    }

    /// <summary>
    /// Tests that Refresh preserves LoadToDataModel settings.
    /// </summary>
    [Fact]
    public async Task Refresh_LoadedToDataModel_PreservesSettings()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_RefreshDM_" + Guid.NewGuid().ToString("N")[..8];

        var mCode = @"let Source = #table({""Val""}, {{42}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create with LoadToDataModel
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToDataModel);

        // Verify initial state
        var loadConfigBefore = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfigBefore.LoadMode);

        // Refresh
        var refreshResult = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(5));
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");

        // Verify LoadToDataModel preserved
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, $"GetLoadConfig after failed: {loadConfigAfter.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfigAfter.LoadMode);

        // Verify still only one table
        var tables = await _dataModelCommands.ListTables(batch);
        Assert.True(tables.Success, $"ListTables failed: {tables.ErrorMessage}");
        var queryTables = tables.Tables.Where(t => t.Name == queryName).ToList();
        Assert.Single(queryTables);
    }

    /// <summary>
    /// Tests that ConnectionOnly load mode is correctly detected.
    /// </summary>
    [Fact]
    public void GetLoadConfig_ConnectionOnly_ReturnsConnectionOnlyMode()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_ConnOnly_" + Guid.NewGuid().ToString("N")[..8];

        var mCode = @"let Source = #table({""Val""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create with ConnectionOnly (no loading)
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Act
        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, queryName);

        // Assert
        Assert.True(loadConfig.Success, $"GetLoadConfig failed: {loadConfig.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, loadConfig.LoadMode);
        Assert.True(string.IsNullOrEmpty(loadConfig.TargetSheet), "ConnectionOnly should not have a target sheet");
    }

    /// <summary>
    /// Tests that LoadToTable mode is correctly detected.
    /// </summary>
    [Fact]
    public void GetLoadConfig_LoadToTable_ReturnsLoadToTableMode()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_Table_" + Guid.NewGuid().ToString("N")[..8];
        var sheetName = "TableSheet";

        var mCode = @"let Source = #table({""Val""}, {{42}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create with LoadToTable
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToTable, sheetName);

        // Act
        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, queryName);

        // Assert
        Assert.True(loadConfig.Success, $"GetLoadConfig failed: {loadConfig.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadConfig.LoadMode);
        Assert.Equal(sheetName, loadConfig.TargetSheet);
    }

    #endregion

    #region LoadToBoth Preservation Tests

    /// <summary>
    /// Tests that Update preserves LoadToBoth settings.
    /// </summary>
    [Fact]
    public async Task Update_LoadedToBoth_PreservesSettings()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_Both_" + Guid.NewGuid().ToString("N")[..8];

        var initialMCode = @"let Source = #table({""A""}, {{1}}) in Source";
        var updatedMCode = @"let Source = #table({""A""}, {{2}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create with LoadToBoth
        _powerQueryCommands.Create(batch, queryName, initialMCode, PowerQueryLoadMode.LoadToBoth);

        // Verify initial state
        var loadConfigBefore = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigBefore.Success, $"GetLoadConfig failed: {loadConfigBefore.ErrorMessage}");
        Assert.True(loadConfigBefore.HasConnection, "Expected HasConnection=true after Create with LoadToBoth");
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, loadConfigBefore.LoadMode);

        // Update
        _powerQueryCommands.Update(batch, queryName, updatedMCode);

        // Verify LoadToBoth preserved
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, $"GetLoadConfig after failed: {loadConfigAfter.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, loadConfigAfter.LoadMode);

        // Verify table exists in Data Model
        var tables = await _dataModelCommands.ListTables(batch);
        Assert.True(tables.Success, $"ListTables failed: {tables.ErrorMessage}");
        var queryTables = tables.Tables.Where(t => t.Name == queryName).ToList();
        Assert.Single(queryTables);
    }

    /// <summary>
    /// Tests that Refresh preserves LoadToBoth settings.
    /// </summary>
    [Fact]
    public async Task Refresh_LoadedToBoth_PreservesSettings()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_RefreshBoth_" + Guid.NewGuid().ToString("N")[..8];

        var mCode = @"let Source = #table({""Val""}, {{99}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create with LoadToBoth
        _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.LoadToBoth);

        // Verify initial state
        var loadConfigBefore = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, loadConfigBefore.LoadMode);

        // Refresh
        var refreshResult = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(5));
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");

        // Verify LoadToBoth preserved
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, $"GetLoadConfig after failed: {loadConfigAfter.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, loadConfigAfter.LoadMode);

        // Verify table still in Data Model
        var tables = await _dataModelCommands.ListTables(batch);
        Assert.True(tables.Success, $"ListTables failed: {tables.ErrorMessage}");
        var queryTables = tables.Tables.Where(t => t.Name == queryName).ToList();
        Assert.Single(queryTables);
    }

    #endregion

    #region Multiple Queries Tests

    /// <summary>
    /// Tests GetLoadConfig correctly identifies different load modes for multiple queries.
    /// </summary>
    [Fact]
    public async Task GetLoadConfig_MultipleQueriesDifferentModes_ReturnsCorrectModeForEach()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var suffix = Guid.NewGuid().ToString("N")[..6];
        var queryConnOnly = "PQ_ConnOnly_" + suffix;
        var queryTable = "PQ_Table_" + suffix;
        var queryDataModel = "PQ_DataModel_" + suffix;
        var queryBoth = "PQ_Both_" + suffix;

        var mCode = @"let Source = #table({""A""}, {{1}}) in Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create queries with different load modes
        _powerQueryCommands.Create(batch, queryConnOnly, mCode, PowerQueryLoadMode.ConnectionOnly);
        _powerQueryCommands.Create(batch, queryTable, mCode, PowerQueryLoadMode.LoadToTable, "Sheet1");
        _powerQueryCommands.Create(batch, queryDataModel, mCode, PowerQueryLoadMode.LoadToDataModel);
        _powerQueryCommands.Create(batch, queryBoth, mCode, PowerQueryLoadMode.LoadToBoth, "Sheet2");

        // Act & Assert - each query should report its correct load mode
        var configConnOnly = _powerQueryCommands.GetLoadConfig(batch, queryConnOnly);
        Assert.True(configConnOnly.Success);
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, configConnOnly.LoadMode);

        var configTable = _powerQueryCommands.GetLoadConfig(batch, queryTable);
        Assert.True(configTable.Success);
        Assert.Equal(PowerQueryLoadMode.LoadToTable, configTable.LoadMode);

        var configDataModel = _powerQueryCommands.GetLoadConfig(batch, queryDataModel);
        Assert.True(configDataModel.Success);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, configDataModel.LoadMode);

        var configBoth = _powerQueryCommands.GetLoadConfig(batch, queryBoth);
        Assert.True(configBoth.Success);
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, configBoth.LoadMode);

        // Verify Data Model has expected tables
        var tables = await _dataModelCommands.ListTables(batch);
        Assert.True(tables.Success);
        Assert.Contains(tables.Tables, t => t.Name == queryDataModel);
        Assert.Contains(tables.Tables, t => t.Name == queryBoth);
        Assert.DoesNotContain(tables.Tables, t => t.Name == queryConnOnly);
        Assert.DoesNotContain(tables.Tables, t => t.Name == queryTable);
    }

    #endregion
}
