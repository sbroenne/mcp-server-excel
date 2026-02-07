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

    #region Schema Change Tests (Column Structure Changes)

    /// <summary>
    /// BUG INVESTIGATION TEST: Column structure changes on Data Model-connected queries.
    ///
    /// Bug Report: Update fails with 0x800A03EC when updating Power Query that:
    /// 1. Is loaded to Data Model
    /// 2. Has schema structure change (add/remove columns)
    ///
    /// This test validates the scenario where:
    /// 1. Create query with 2 columns loaded to Data Model
    /// 2. Update query to add a 3rd column
    /// 3. Expected: Either succeeds OR throws meaningful exception
    /// </summary>
    [Fact]
    public async Task Update_LoadedToDataModel_AddColumn_HandlesSchemaChange()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_SchemaChange_" + Guid.NewGuid().ToString("N")[..8];

        // Initial M code with 2 columns
        var twoColumnMCode = @"let
    Source = #table(
        {""ID"", ""Name""},
        {
            {1, ""Alpha""},
            {2, ""Beta""}
        }
    )
in
    Source";

        // Updated M code with 3 columns (ADDS ""Amount"" column)
        var threeColumnMCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Amount""},
        {
            {1, ""Alpha"", 100},
            {2, ""Beta"", 200},
            {3, ""Gamma"", 300}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery and load to Data Model
        _powerQueryCommands.Create(batch, queryName, twoColumnMCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Verify initial state - 1 table with 2 columns
        var tablesBefore = await _dataModelCommands.ListTables(batch);
        Assert.True(tablesBefore.Success, $"ListTables before failed: {tablesBefore.ErrorMessage}");
        Assert.Single(tablesBefore.Tables, t => t.Name == queryName);

        var tableBefore = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableBefore.Success, $"ReadTable before failed: {tableBefore.ErrorMessage}");
        Assert.Equal(2, tableBefore.Columns.Count);  // ID, Name

        // STEP 3: Update the M code to ADD A COLUMN
        // This is the bug scenario - schema change on Data Model-connected query
        _powerQueryCommands.Update(batch, queryName, threeColumnMCode);

        // STEP 4: Verify load configuration is PRESERVED
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, $"GetLoadConfig after failed: {loadConfigAfter.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfigAfter.LoadMode);

        // STEP 5: Verify Data Model table now has 3 columns
        var tableAfter = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableAfter.Success, $"ReadTable after failed: {tableAfter.ErrorMessage}");
        Assert.Equal(3, tableAfter.Columns.Count);  // ID, Name, Amount

        // STEP 6: Verify M code was actually updated
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("Amount", viewResult.MCode);
    }

    /// <summary>
    /// BUG INVESTIGATION TEST: Column removal on Data Model-connected queries.
    ///
    /// This test validates the scenario where:
    /// 1. Create query with 3 columns loaded to Data Model
    /// 2. Update query to remove a column (now 2 columns)
    /// 3. Expected: Either succeeds OR throws meaningful exception
    /// </summary>
    [Fact]
    public async Task Update_LoadedToDataModel_RemoveColumn_HandlesSchemaChange()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_RemoveCol_" + Guid.NewGuid().ToString("N")[..8];

        // Initial M code with 3 columns
        var threeColumnMCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Amount""},
        {
            {1, ""Alpha"", 100},
            {2, ""Beta"", 200}
        }
    )
in
    Source";

        // Updated M code with 2 columns (REMOVES ""Amount"" column)
        var twoColumnMCode = @"let
    Source = #table(
        {""ID"", ""Name""},
        {
            {1, ""Alpha""},
            {2, ""Beta""},
            {3, ""Gamma""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery and load to Data Model
        _powerQueryCommands.Create(batch, queryName, threeColumnMCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Verify initial state - 1 table with 3 columns
        var tableBefore = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableBefore.Success, $"ReadTable before failed: {tableBefore.ErrorMessage}");
        Assert.Equal(3, tableBefore.Columns.Count);  // ID, Name, Amount

        // STEP 3: Update the M code to REMOVE A COLUMN
        _powerQueryCommands.Update(batch, queryName, twoColumnMCode);

        // STEP 4: Verify load configuration is PRESERVED
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, $"GetLoadConfig after failed: {loadConfigAfter.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfigAfter.LoadMode);

        // STEP 5: Verify Data Model table now has 2 columns
        var tableAfter = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableAfter.Success, $"ReadTable after failed: {tableAfter.ErrorMessage}");
        Assert.Equal(2, tableAfter.Columns.Count);  // ID, Name

        // STEP 6: Verify M code was actually updated
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.DoesNotContain("Amount", viewResult.MCode);
    }

    /// <summary>
    /// BUG INVESTIGATION TEST: Column type change on Data Model-connected queries.
    ///
    /// This test validates the scenario where:
    /// 1. Create query with columns loaded to Data Model
    /// 2. Update query to change column data type (e.g., number to text)
    /// 3. Expected: Either succeeds OR throws meaningful exception
    /// </summary>
    [Fact]
    public async Task Update_LoadedToDataModel_ChangeColumnType_HandlesSchemaChange()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_TypeChange_" + Guid.NewGuid().ToString("N")[..8];

        // Initial M code with Amount as number
        var numberTypeMCode = @"let
    Source = #table(
        {""ID"", ""Amount""},
        {
            {1, 100},
            {2, 200}
        }
    )
in
    Source";

        // Updated M code with Amount as text
        var textTypeMCode = @"let
    Source = #table(
        {""ID"", ""Amount""},
        {
            {1, ""One Hundred""},
            {2, ""Two Hundred""},
            {3, ""Three Hundred""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery and load to Data Model
        _powerQueryCommands.Create(batch, queryName, numberTypeMCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Verify initial state
        var tableBefore = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableBefore.Success, $"ReadTable before failed: {tableBefore.ErrorMessage}");
        Assert.Equal(2, tableBefore.Columns.Count);

        // STEP 3: Update the M code to CHANGE COLUMN TYPE
        _powerQueryCommands.Update(batch, queryName, textTypeMCode);

        // STEP 4: Verify load configuration is PRESERVED
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, $"GetLoadConfig after failed: {loadConfigAfter.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfigAfter.LoadMode);

        // STEP 5: Verify M code was actually updated
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("One Hundred", viewResult.MCode);
    }

    /// <summary>
    /// BUG INVESTIGATION TEST: Schema change when DAX measure references the table.
    ///
    /// Bug Report: Update fails with 0x800A03EC for Power Query AND 0x800AC472 for DAX measures
    /// when updating Power Query that has DAX measures referencing it.
    ///
    /// This test validates the scenario where:
    /// 1. Create query with columns loaded to Data Model
    /// 2. Create DAX measure that references a column
    /// 3. Update query to add a column (schema change)
    /// 4. Expected: Either succeeds OR throws meaningful exception
    /// </summary>
    [Fact]
    public async Task Update_LoadedToDataModel_WithDaxMeasure_AddColumn_HandlesSchemaChange()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_DaxMeasure_" + Guid.NewGuid().ToString("N")[..8];

        // Initial M code with 2 columns
        var twoColumnMCode = @"let
    Source = #table(
        {""ID"", ""Amount""},
        {
            {1, 100},
            {2, 200},
            {3, 300}
        }
    )
in
    Source";

        // Updated M code with 3 columns (ADDS ""Category"" column)
        var threeColumnMCode = @"let
    Source = #table(
        {""ID"", ""Amount"", ""Category""},
        {
            {1, 100, ""A""},
            {2, 200, ""B""},
            {3, 300, ""A""},
            {4, 400, ""C""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery and load to Data Model
        _powerQueryCommands.Create(batch, queryName, twoColumnMCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Verify table exists in Data Model
        var tableBefore = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableBefore.Success, $"ReadTable before failed: {tableBefore.ErrorMessage}");
        Assert.Equal(2, tableBefore.Columns.Count);  // ID, Amount

        // STEP 3: Create DAX measure that references the Amount column
        var measureName = "TotalAmount";
        var daxFormula = $"SUM('{queryName}'[Amount])";
        _dataModelCommands.CreateMeasure(batch, queryName, measureName, daxFormula);  // Throws on error

        // STEP 4: Verify measure exists
        var measuresBefore = _dataModelCommands.ListMeasures(batch);
        Assert.True(measuresBefore.Success, $"ListMeasures failed: {measuresBefore.ErrorMessage}");
        Assert.Contains(measuresBefore.Measures, m => m.Name == measureName);

        // STEP 5: Update the M code to ADD A COLUMN (this is the bug scenario)
        // The DAX measure references Amount - adding Category should NOT break it
        _powerQueryCommands.Update(batch, queryName, threeColumnMCode);

        // STEP 6: Verify load configuration is PRESERVED
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, $"GetLoadConfig after failed: {loadConfigAfter.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfigAfter.LoadMode);

        // STEP 7: Verify Data Model table now has 3 columns
        var tableAfter = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableAfter.Success, $"ReadTable after failed: {tableAfter.ErrorMessage}");
        Assert.Equal(3, tableAfter.Columns.Count);  // ID, Amount, Category

        // STEP 8: Verify DAX measure STILL EXISTS and is valid
        var measuresAfter = _dataModelCommands.ListMeasures(batch);
        Assert.True(measuresAfter.Success, $"ListMeasures after failed: {measuresAfter.ErrorMessage}");
        Assert.Contains(measuresAfter.Measures, m => m.Name == measureName);

        // STEP 9: Verify M code was actually updated
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("Category", viewResult.MCode);
    }

    /// <summary>
    /// BUG INVESTIGATION TEST: Schema change removes column referenced by DAX measure.
    ///
    /// This is the MOST DANGEROUS scenario - removing a column that a DAX measure depends on.
    ///
    /// This test validates the scenario where:
    /// 1. Create query with columns loaded to Data Model
    /// 2. Create DAX measure that references a specific column
    /// 3. Update query to REMOVE that column
    /// 4. Expected: Should fail gracefully with meaningful error (DAX measure becomes invalid)
    /// </summary>
    [Fact]
    public async Task Update_LoadedToDataModel_WithDaxMeasure_RemoveReferencedColumn_HandlesError()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_DaxBreak_" + Guid.NewGuid().ToString("N")[..8];

        // Initial M code with Amount column
        var withAmountMCode = @"let
    Source = #table(
        {""ID"", ""Name"", ""Amount""},
        {
            {1, ""Alpha"", 100},
            {2, ""Beta"", 200}
        }
    )
in
    Source";

        // Updated M code WITHOUT Amount column (REMOVES column that DAX measure references)
        var withoutAmountMCode = @"let
    Source = #table(
        {""ID"", ""Name""},
        {
            {1, ""Alpha""},
            {2, ""Beta""},
            {3, ""Gamma""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery and load to Data Model
        _powerQueryCommands.Create(batch, queryName, withAmountMCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Verify table exists with Amount column
        var tableBefore = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableBefore.Success, $"ReadTable before failed: {tableBefore.ErrorMessage}");
        Assert.Equal(3, tableBefore.Columns.Count);  // ID, Name, Amount

        // STEP 3: Create DAX measure that references the Amount column
        var measureName = "TotalAmount";
        var daxFormula = $"SUM('{queryName}'[Amount])";
        _dataModelCommands.CreateMeasure(batch, queryName, measureName, daxFormula);  // Throws on error

        // STEP 4: Update the M code to REMOVE the Amount column
        // This should cause issues because the DAX measure references Amount
        // The system should either:
        // a) Fail gracefully with a meaningful error, OR
        // b) Succeed but leave the DAX measure in an invalid state
        _powerQueryCommands.Update(batch, queryName, withoutAmountMCode);

        // STEP 5: Verify M code was updated (the update itself should succeed)
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.DoesNotContain("Amount", viewResult.MCode);

        // STEP 6: Verify table no longer has Amount column
        var tableAfter = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableAfter.Success, $"ReadTable after failed: {tableAfter.ErrorMessage}");
        Assert.Equal(2, tableAfter.Columns.Count);  // ID, Name only

        // STEP 7: The DAX measure may still exist but reference an invalid column
        // This is expected behavior - Excel/Power Pivot doesn't auto-delete measures
        var measuresAfter = _dataModelCommands.ListMeasures(batch);
        Assert.True(measuresAfter.Success, $"ListMeasures after failed: {measuresAfter.ErrorMessage}");
        // Measure may or may not still exist - either is acceptable
    }

    /// <summary>
    /// BUG INVESTIGATION TEST: Update DAX measure after schema change.
    ///
    /// Bug Report: DAX measure update fails with 0x800AC472 after Power Query update.
    ///
    /// This test validates the scenario where:
    /// 1. Create query loaded to Data Model
    /// 2. Create DAX measure
    /// 3. Update Power Query (schema change)
    /// 4. Update the DAX measure formula
    /// 5. Expected: Either succeeds OR throws meaningful exception
    /// </summary>
    [Fact]
    public async Task Update_DaxMeasure_AfterSchemaChange_HandlesUpdate()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_DaxUpdate_" + Guid.NewGuid().ToString("N")[..8];

        // Initial M code
        var initialMCode = @"let
    Source = #table(
        {""ID"", ""Amount""},
        {
            {1, 100},
            {2, 200}
        }
    )
in
    Source";

        // Updated M code with new column
        var updatedMCode = @"let
    Source = #table(
        {""ID"", ""Amount"", ""Quantity""},
        {
            {1, 100, 5},
            {2, 200, 10},
            {3, 300, 15}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery and load to Data Model
        _powerQueryCommands.Create(batch, queryName, initialMCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Create initial DAX measure
        var measureName = "TotalAmount";
        var initialDaxFormula = $"SUM('{queryName}'[Amount])";
        _dataModelCommands.CreateMeasure(batch, queryName, measureName, initialDaxFormula);  // Throws on error

        // STEP 3: Update the M code (schema change - adds Quantity column)
        _powerQueryCommands.Update(batch, queryName, updatedMCode);

        // STEP 4: Verify schema change worked
        var tableAfter = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableAfter.Success, $"ReadTable after schema change failed: {tableAfter.ErrorMessage}");
        Assert.Equal(3, tableAfter.Columns.Count);  // ID, Amount, Quantity

        // STEP 5: Update the DAX measure to use the NEW column (this is the 0x800AC472 bug scenario)
        var updatedDaxFormula = $"SUM('{queryName}'[Amount]) + SUM('{queryName}'[Quantity])";
        _dataModelCommands.UpdateMeasure(batch, measureName, updatedDaxFormula);  // Throws on error

        // STEP 6: Verify measure was updated
        var readMeasure = _dataModelCommands.Read(batch, measureName);
        Assert.True(readMeasure.Success, $"Read measure failed: {readMeasure.ErrorMessage}");
        Assert.Contains("Quantity", readMeasure.DaxFormula);
    }

    /// <summary>
    /// BUG INVESTIGATION TEST: Multiple sequential schema changes with DAX measures.
    ///
    /// This tests the compounding effect of multiple updates, which may trigger
    /// the bug more reliably than a single update.
    /// </summary>
    [Fact]
    public async Task Update_LoadedToDataModel_MultipleSchemaChanges_WithDaxMeasures()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_MultiChange_" + Guid.NewGuid().ToString("N")[..8];

        // Version 1: 2 columns
        var v1MCode = @"let
    Source = #table(
        {""ID"", ""Value""},
        {
            {1, 100},
            {2, 200}
        }
    )
in
    Source";

        // Version 2: 3 columns (add Category)
        var v2MCode = @"let
    Source = #table(
        {""ID"", ""Value"", ""Category""},
        {
            {1, 100, ""A""},
            {2, 200, ""B""}
        }
    )
in
    Source";

        // Version 3: 4 columns (add Quantity)
        var v3MCode = @"let
    Source = #table(
        {""ID"", ""Value"", ""Category"", ""Quantity""},
        {
            {1, 100, ""A"", 5},
            {2, 200, ""B"", 10},
            {3, 300, ""C"", 15}
        }
    )
in
    Source";

        // Version 4: Back to 3 columns (remove Category)
        var v4MCode = @"let
    Source = #table(
        {""ID"", ""Value"", ""Quantity""},
        {
            {1, 100, 5},
            {2, 200, 10},
            {3, 300, 15},
            {4, 400, 20}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery V1
        _powerQueryCommands.Create(batch, queryName, v1MCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Create DAX measure on Value column
        var measureName = "SumValue";
        var daxFormula = $"SUM('{queryName}'[Value])";
        _dataModelCommands.CreateMeasure(batch, queryName, measureName, daxFormula);  // Throws on error

        // STEP 3: Update to V2 (add Category)
        _powerQueryCommands.Update(batch, queryName, v2MCode);
        var tableV2 = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableV2.Success, $"ReadTable V2 failed: {tableV2.ErrorMessage}");
        Assert.Equal(3, tableV2.Columns.Count);

        // STEP 4: Update to V3 (add Quantity)
        _powerQueryCommands.Update(batch, queryName, v3MCode);
        var tableV3 = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableV3.Success, $"ReadTable V3 failed: {tableV3.ErrorMessage}");
        Assert.Equal(4, tableV3.Columns.Count);

        // STEP 5: Update to V4 (remove Category)
        _powerQueryCommands.Update(batch, queryName, v4MCode);
        var tableV4 = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableV4.Success, $"ReadTable V4 failed: {tableV4.ErrorMessage}");
        Assert.Equal(3, tableV4.Columns.Count);  // ID, Value, Quantity

        // STEP 6: Verify DAX measure still exists (it references Value which was never removed)
        var measuresAfter = _dataModelCommands.ListMeasures(batch);
        Assert.True(measuresAfter.Success, $"ListMeasures after failed: {measuresAfter.ErrorMessage}");
        Assert.Contains(measuresAfter.Measures, m => m.Name == measureName);

        // STEP 7: Verify load mode preserved after all changes
        var loadConfig = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfig.Success, $"GetLoadConfig failed: {loadConfig.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, loadConfig.LoadMode);
    }

    /// <summary>
    /// BUG INVESTIGATION TEST: Schema change with complex M code transformations.
    ///
    /// Tests with more realistic M code that includes transformations,
    /// not just simple #table definitions.
    /// </summary>
    [Fact]
    public async Task Update_LoadedToDataModel_ComplexMCode_SchemaChange()
    {
        // Arrange
        var testExcelFile = _fixture.CreateTestFile();
        var queryName = "PQ_Complex_" + Guid.NewGuid().ToString("N")[..8];

        // Initial M code with transformations
        var initialMCode = @"let
    Source = #table(
        {""ID"", ""RawValue""},
        {
            {1, 100},
            {2, 200},
            {3, 300}
        }
    ),
    AddedColumn = Table.AddColumn(Source, ""DoubleValue"", each [RawValue] * 2),
    ChangedType = Table.TransformColumnTypes(AddedColumn, {{""DoubleValue"", type number}})
in
    ChangedType";

        // Updated M code with additional transformation step
        var updatedMCode = @"let
    Source = #table(
        {""ID"", ""RawValue""},
        {
            {1, 100},
            {2, 200},
            {3, 300},
            {4, 400}
        }
    ),
    AddedColumn = Table.AddColumn(Source, ""DoubleValue"", each [RawValue] * 2),
    AddedTriple = Table.AddColumn(AddedColumn, ""TripleValue"", each [RawValue] * 3),
    ChangedType = Table.TransformColumnTypes(AddedTriple, {{""DoubleValue"", type number}, {""TripleValue"", type number}})
in
    ChangedType";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Create PowerQuery with complex M code
        _powerQueryCommands.Create(batch, queryName, initialMCode, PowerQueryLoadMode.LoadToDataModel);

        // STEP 2: Verify initial state - should have 3 columns (ID, RawValue, DoubleValue)
        var tableBefore = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableBefore.Success, $"ReadTable before failed: {tableBefore.ErrorMessage}");
        Assert.Equal(3, tableBefore.Columns.Count);

        // STEP 3: Create DAX measure
        var measureName = "AvgDouble";
        var daxFormula = $"AVERAGE('{queryName}'[DoubleValue])";
        _dataModelCommands.CreateMeasure(batch, queryName, measureName, daxFormula);  // Throws on error

        // STEP 4: Update with more complex M code (adds TripleValue column)
        _powerQueryCommands.Update(batch, queryName, updatedMCode);

        // STEP 5: Verify schema change - should now have 4 columns
        var tableAfter = await _dataModelCommands.ReadTable(batch, queryName);
        Assert.True(tableAfter.Success, $"ReadTable after failed: {tableAfter.ErrorMessage}");
        Assert.Equal(4, tableAfter.Columns.Count);  // ID, RawValue, DoubleValue, TripleValue

        // STEP 6: Verify DAX measure still valid
        var measuresAfter = _dataModelCommands.ListMeasures(batch);
        Assert.True(measuresAfter.Success, $"ListMeasures after failed: {measuresAfter.ErrorMessage}");
        Assert.Contains(measuresAfter.Measures, m => m.Name == measureName);

        // STEP 7: Verify M code was updated
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");
        Assert.Contains("TripleValue", viewResult.MCode);
    }

    #endregion
}




