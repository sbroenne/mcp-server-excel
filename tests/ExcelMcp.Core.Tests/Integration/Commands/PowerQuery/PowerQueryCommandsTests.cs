using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Integration tests for Power Query operations focusing on LLM use cases.
/// Tests cover the essential workflows: import, list, view, update, delete, refresh with load destinations.
/// Uses PowerQueryTestsFixture which creates ONE Power Query file per test class.
/// Each test uses unique files for complete isolation where needed.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
[Trait("Speed", "Medium")]
public partial class PowerQueryCommandsTests : IClassFixture<PowerQueryTestsFixture>
{
    private readonly PowerQueryCommands _powerQueryCommands;
    private readonly SheetCommands _sheetCommands;
    private readonly string _powerQueryFile;
    private readonly PowerQueryCreationResult _creationResult;
    private readonly string _tempDir;

    /// <summary>
    /// Initializes a new instance of the <see cref="PowerQueryCommandsTests"/> class.
    /// </summary>
    public PowerQueryCommandsTests(PowerQueryTestsFixture fixture)
    {
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
        _sheetCommands = new SheetCommands();
        _powerQueryFile = fixture.TestFilePath;
        _creationResult = fixture.CreationResult;
        _tempDir = Path.GetDirectoryName(fixture.TestFilePath)!;
    }

    #region Core Lifecycle Tests (6 tests)

    /// <summary>
    /// Validates that the fixture creation succeeded (import operation).
    /// LLM use case: "import a Power Query from a .pq file"
    /// </summary>
    [Fact]
    public void Import_ViaFixture_CreatesQueriesSuccessfully()
    {
        // Assert the fixture creation succeeded
        Assert.True(_creationResult.Success,
            $"Power Query creation failed during fixture initialization: {_creationResult.ErrorMessage}");

        Assert.True(_creationResult.FileCreated, "File creation failed");
        Assert.Equal(3, _creationResult.MCodeFilesCreated);
        Assert.Equal(3, _creationResult.QueriesImported);
        Assert.True(_creationResult.CreationTimeSeconds > 0);
    }

    /// <summary>
    /// Tests basic import operation with M code file.
    /// LLM use case: "import this M code as a new Power Query"
    /// </summary>
    [Fact]
    public void Import_ValidMCode_ReturnsSuccess()
    {
        // Arrange - Use unique file to avoid polluting fixture
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Import_ValidMCode_ReturnsSuccess),
            _tempDir);
        var queryName = "TestQuery";
        var mCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        // Act
        using var batch = ExcelSession.BeginBatch(testExcelFile);
        var result = _powerQueryCommands.Create(batch, queryName, mCode, PowerQueryLoadMode.ConnectionOnly);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    /// <summary>
    /// Tests listing queries in a workbook.
    /// LLM use case: "show me all Power Queries in this workbook"
    /// </summary>
    [Fact]
    public void List_FixtureWorkbook_ReturnsFixtureQueries()
    {
        // Act
        using var batch = ExcelSession.BeginBatch(_powerQueryFile);
        var result = _powerQueryCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);
        Assert.Equal(3, result.Queries.Count);
    }

    /// <summary>
    /// Tests viewing M code of an existing query.
    /// LLM use case: "show me the M code for this query"
    /// </summary>
    [Fact]
    public void View_BasicQuery_ReturnsMCode()
    {
        // Act
        using var batch = ExcelSession.BeginBatch(_powerQueryFile);
        var result = _powerQueryCommands.View(batch, "BasicQuery");

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.MCode);
        Assert.Contains("Source", result.MCode);
    }

    /// <summary>
    /// Tests updating existing query with new M code.
    /// LLM use case: "update this query's M code"
    /// </summary>
    [Fact]
    public void Update_ExistingQuery_ReturnsSuccess()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Update_ExistingQuery_ReturnsSuccess),
            _tempDir);

        var queryName = "PQ_Update_" + Guid.NewGuid().ToString("N")[..8];
        var originalMCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";
        var updatedMCode = @"let
    UpdatedSource = 1
in
    UpdatedSource";

        // Act
        using var batch = ExcelSession.BeginBatch(testExcelFile);
        _powerQueryCommands.Create(batch, queryName, originalMCode);
        var result = _powerQueryCommands.Update(batch, queryName, updatedMCode);

        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// REGRESSION TEST for bug report: Update action merges instead of replaces M code
    ///
    /// Bug: Update was concatenating/merging new M code with existing M code instead of replacing it,
    /// resulting in severely corrupted Power Query definitions with triple-merged comments, multiple let blocks,
    /// and invalid M syntax.
    ///
    /// This test validates that Update completely REPLACES M code (not merges/appends).
    /// LLM use case: "update this query's M code and verify it was replaced"
    /// </summary>
    [Fact]
    public void Update_ExistingQuery_ReplacesNotMergesMCode()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Update_ExistingQuery_ReplacesNotMergesMCode),
            _tempDir);

        var queryName = "PQ_ReplaceTest_" + Guid.NewGuid().ToString("N")[..8];

        // Original M code with distinctive markers
        var originalMCode = @"let
    OriginalSource = ""ORIGINAL_MARKER"",
    OriginalStep = ""Should be completely removed""
in
    OriginalSource";

        // New M code that should completely replace original
        var newMCode = @"let
    NewSource = ""NEW_MARKER"",
    NewStep = ""Should be the only content""
in
    NewSource";

        // Act
        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Step 1: Create query with original M code
        var createResult = _powerQueryCommands.Create(batch, queryName, originalMCode);
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");

        // Step 2: Update with new M code
        var updateResult = _powerQueryCommands.Update(batch, queryName, newMCode);
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // Step 3: View the resulting M code
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success, $"View failed: {viewResult.ErrorMessage}");

        // Assert - CRITICAL: Verify M code was REPLACED, not merged
        // 1. Should contain the new M code
        Assert.Contains("NEW_MARKER", viewResult.MCode);
        Assert.Contains("NewSource", viewResult.MCode);
        Assert.Contains("Should be the only content", viewResult.MCode);

        // 2. Should NOT contain any traces of the original M code
        Assert.DoesNotContain("ORIGINAL_MARKER", viewResult.MCode);
        Assert.DoesNotContain("OriginalSource", viewResult.MCode);
        Assert.DoesNotContain("Should be completely removed", viewResult.MCode);

        // 3. Should not have duplicate 'let' or 'in' keywords (sign of merging)
        int letCount = System.Text.RegularExpressions.Regex.Matches(viewResult.MCode, @"\blet\b").Count;
        int inCount = System.Text.RegularExpressions.Regex.Matches(viewResult.MCode, @"\bin\b").Count;
        Assert.Equal(1, letCount);
        Assert.Equal(1, inCount);
    }

    /// <summary>
    /// REGRESSION TEST: Multiple sequential updates should each completely replace M code
    ///
    /// This test validates that the merging bug doesn't compound with multiple updates.
    /// LLM use case: "update this query multiple times during development"
    /// </summary>
    [Fact]
    public void Update_MultipleSequentialUpdates_EachReplacesCompletely()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Update_MultipleSequentialUpdates_EachReplacesCompletely),
            _tempDir);

        var queryName = "PQ_MultiUpdate_" + Guid.NewGuid().ToString("N")[..8];

        // Create three different M code versions
        var version1MCode = @"let
    V1 = ""VERSION_1""
in
    V1";

        var version2MCode = @"let
    V2 = ""VERSION_2""
in
    V2";

        var version3MCode = @"let
    V3 = ""VERSION_3""
in
    V3";

        // Act
        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Create with version 1
        _powerQueryCommands.Create(batch, queryName, version1MCode);

        // Update to version 2
        _powerQueryCommands.Update(batch, queryName, version2MCode);

        // Update to version 3
        _powerQueryCommands.Update(batch, queryName, version3MCode);

        // View final result
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success);

        // Assert - Should only have version 3, no traces of v1 or v2
        Assert.Contains("VERSION_3", viewResult.MCode);
        Assert.DoesNotContain("VERSION_1", viewResult.MCode);
        Assert.DoesNotContain("VERSION_2", viewResult.MCode);

        // Verify no compound merging (should still have exactly 1 let/in)
        int letCount = System.Text.RegularExpressions.Regex.Matches(viewResult.MCode, @"\blet\b").Count;
        int inCount = System.Text.RegularExpressions.Regex.Matches(viewResult.MCode, @"\bin\b").Count;
        Assert.Equal(1, letCount);
        Assert.Equal(1, inCount);
    }

    /// <summary>
    /// Tests deleting an existing query.
    /// LLM use case: "delete this Power Query"
    /// </summary>
    [Fact]
    public void Delete_ExistingQuery_ReturnsSuccess()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Delete_ExistingQuery_ReturnsSuccess),
            _tempDir);

        var queryName = "PQ_Delete_" + Guid.NewGuid().ToString("N")[..8];
        var mCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        // Act
        using var batch = ExcelSession.BeginBatch(testExcelFile);
        _powerQueryCommands.Create(batch, queryName, mCode);
        var result = _powerQueryCommands.Delete(batch, queryName);

        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// Tests that attempting to Create a query that already exists returns an error.
    /// LLM use case: "accidentally trying to create the same query twice"
    /// Real bug: LLM using Create action on existing query instead of Update
    /// </summary>
    [Fact]
    public void Create_DuplicateQueryName_ReturnsError()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Create_DuplicateQueryName_ReturnsError),
            _tempDir);

        var queryName = "TestQuery";
        var mCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Act 1: Create query first time (should succeed)
        var firstCreate = _powerQueryCommands.Create(batch, queryName, mCode);
        Assert.True(firstCreate.Success, $"First create should succeed: {firstCreate.ErrorMessage}");

        // Act 2 & Assert: Try to Create same query again (should throw InvalidOperationException)
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _powerQueryCommands.Create(batch, queryName, mCode));

        Assert.Contains("already exists", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(queryName, exception.Message);

        // Verify query still exists and wasn't corrupted
        var viewResult = _powerQueryCommands.View(batch, queryName);
        Assert.True(viewResult.Success);
        Assert.NotEmpty(viewResult.MCode);
    }

    #endregion

    #region Load Destination Workflows (3 tests)

    #endregion

    #region Advanced Use Cases (1 test)

    /// <summary>
    /// Tests that one Power Query can successfully reference another Power Query.
    /// LLM use case: "create a query that filters data from another query"
    /// </summary>
    [Fact]
    public void Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully),
            _tempDir);

        // Create M code for the source query (base data)
        string sourceQueryMCode = @"let
    Source = #table(
        {""ProductID"", ""ProductName"", ""Price""},
        {
            {1, ""Widget"", 10.99},
            {2, ""Gadget"", 25.50},
            {3, ""Doohickey"", 15.75}
        }
    )
in
    Source";

        // Create M code for the derived query (references the source query)
        string derivedQueryMCode = @"let
    Source = SourceQuery,
    FilteredRows = Table.SelectRows(Source, each [Price] > 15)
in
    FilteredRows";

        // Act & Assert
        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Import source query first
        var sourceImportResult = _powerQueryCommands.Create(
            batch,
            "SourceQuery",
            sourceQueryMCode,
            loadMode: PowerQueryLoadMode.LoadToTable);

        Assert.True(sourceImportResult.Success,
            $"Source query import failed: {sourceImportResult.ErrorMessage}");

        // Import derived query (references SourceQuery)
        var derivedImportResult = _powerQueryCommands.Create(
            batch,
            "DerivedQuery",
            derivedQueryMCode,
            loadMode: PowerQueryLoadMode.LoadToTable);

        Assert.True(derivedImportResult.Success,
            $"Derived query import failed: {derivedImportResult.ErrorMessage}");

        // Verify both queries exist in the workbook
        var listResult = _powerQueryCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Equal(2, listResult.Queries.Count);
        Assert.Contains(listResult.Queries, q => q.Name == "SourceQuery");
        Assert.Contains(listResult.Queries, q => q.Name == "DerivedQuery");

        // Verify the derived query M code references SourceQuery
        var derivedViewResult = _powerQueryCommands.View(batch, "DerivedQuery");
        Assert.True(derivedViewResult.Success);
        Assert.Contains("SourceQuery", derivedViewResult.MCode);
        Assert.Contains("Table.SelectRows", derivedViewResult.MCode);

        // Refresh both queries to ensure they execute successfully
        var sourceRefreshResult = _powerQueryCommands.Refresh(batch, "SourceQuery", TimeSpan.FromMinutes(5));
        Assert.True(sourceRefreshResult.Success,
            $"Source query refresh failed: {sourceRefreshResult.ErrorMessage}");

        var derivedRefreshResult = _powerQueryCommands.Refresh(batch, "DerivedQuery", TimeSpan.FromMinutes(5));
        Assert.True(derivedRefreshResult.Success,
            $"Derived query refresh failed: {derivedRefreshResult.ErrorMessage}");
    }

    #endregion

    #region Regression Tests

    /// <summary>
    /// REGRESSION TEST for reported LLM bug:
    /// 1. Create PowerQuery that loads to sheet - works
    /// 2. Update the query and run again
    /// 3. Query turns into connection-only (BUG!)
    ///
    /// This test verifies that Update preserves the load configuration.
    /// </summary>
    [Fact]
    public void Update_QueryLoadedToSheet_PreservesLoadConfiguration()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Update_QueryLoadedToSheet_PreservesLoadConfiguration),
            _tempDir);

        var queryName = "LoadedQuery_" + Guid.NewGuid().ToString("N")[..8];
        var sheetName = "DataSheet";

        // Initial M code for the query
        string initialMCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        // Updated M code for the query
        string updatedMCode = @"let
    UpdatedSource = #table(
        {""NewCol1"", ""NewCol2""},
        {
            {""Updated1"", ""Updated2""},
            {""Data1"", ""Data2""}
        }
    )
in
    UpdatedSource";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // STEP 1: Import query and load to worksheet
        var importResult = _powerQueryCommands.Create(batch, queryName, initialMCode, PowerQueryLoadMode.LoadToTable, sheetName);
        Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

        // Verify initial load configuration
        var loadConfigBefore = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigBefore.Success, "GetLoadConfig before update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadConfigBefore.LoadMode);
        Assert.Equal(sheetName, loadConfigBefore.TargetSheet);

        // STEP 2: Update the query M code (now auto-refreshes)
        var updateResult = _powerQueryCommands.Update(batch, queryName, updatedMCode);
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // STEP 3: Verify load configuration is PRESERVED (regression check)
        var loadConfigAfter = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfter.Success, "GetLoadConfig after update failed");

        // THE BUG: This assertion should pass but might fail if Update doesn't restore load config
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadConfigAfter.LoadMode);
        Assert.Equal(sheetName, loadConfigAfter.TargetSheet);

        // STEP 4: Verify data is still loaded to the sheet (not connection-only)
        var listResult = _sheetCommands.List(batch);
        Assert.Contains(listResult.Worksheets, s => s.Name == sheetName);
    }

    /// <summary>
    /// REGRESSION TEST for reported user bug (2025-01-28):
    /// User workflow: Create query loaded to worksheet ? UpdateMCode ? Refresh ? query becomes connection-only
    ///
    /// This test validates that UpdateMCode + Refresh preserves load configuration.
    /// Expected: Load configuration should survive both UpdateMCode AND Refresh operations.
    /// </summary>
    [Fact]
    public void UpdateMCodeThenRefresh_QueryLoadedToSheet_PreservesLoadConfiguration()
    {
        // Arrange

        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(UpdateMCodeThenRefresh_QueryLoadedToSheet_PreservesLoadConfiguration),
            _tempDir);

        var queryName = "LoadedQuery_" + Guid.NewGuid().ToString("N")[..8];
        var sheetName = "DataSheet";

        // Initial M code for the query
        string initialMCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        // Updated M code for the query
        string updatedMCode = @"let
    UpdatedSource = #table(
        {""NewCol1"", ""NewCol2""},
        {
            {""Updated1"", ""Updated2""},
            {""Data1"", ""Data2""}
        }
    )
in
    UpdatedSource";

        using var batch = ExcelSession.BeginBatch(testFile);

        // STEP 1: Create query and load to worksheet
        var createResult = _powerQueryCommands.Create(batch, queryName, initialMCode, PowerQueryLoadMode.LoadToTable, sheetName);
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");

        // STEP 2: Verify initial load configuration
        var loadConfigBefore = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigBefore.Success, "GetLoadConfig before update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadConfigBefore.LoadMode);
        Assert.Equal(sheetName, loadConfigBefore.TargetSheet);

        // STEP 3: Update M code (now auto-refreshes - this is the simplified API)
        var updateResult = _powerQueryCommands.Update(batch, queryName, updatedMCode);
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // STEP 4: THE CRITICAL CHECK - Does load config survive Update (which includes refresh)?
        var loadConfigAfterUpdate = _powerQueryCommands.GetLoadConfig(batch, queryName);
        Assert.True(loadConfigAfterUpdate.Success, "GetLoadConfig after update failed");

        // This assertion verifies load config is preserved through update+refresh
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadConfigAfterUpdate.LoadMode);
        Assert.Equal(sheetName, loadConfigAfterUpdate.TargetSheet);

        // STEP 5: Verify data is actually on the sheet (not connection-only)

        Assert.False(string.IsNullOrEmpty(loadConfigAfterUpdate.TargetSheet),
            "Query should have a target sheet (not be connection-only)");
    }

    #endregion

    #region Column Structure Regression Tests (2 tests)

    /// <summary>
    /// REGRESSION TEST: Validates Update properly handles column structure changes
    ///
    /// LLM use case: "update a query to change column structure and verify columns update"
    ///
    /// Scenario:
    /// 1. Create a PowerQuery with one column and load to a worksheet
    /// 2. Check that there is only one column
    /// 3. Update the query and load again
    /// 4. Check that there is still only one column
    /// 5. Update the query to create two columns and load again
    /// 6. Check that there are two columns (validates column structure updates correctly)
    ///
    /// Historical bug: QueryTable.PreserveColumnInfo=true prevented column updates
    /// Fix: Set PreserveColumnInfo=false and clear worksheet before recreating QueryTable
    /// </summary>
    [Fact]
    public void Update_QueryColumnStructure_UpdatesWorksheetColumns()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Update_QueryColumnStructure_UpdatesWorksheetColumns),
            _tempDir);

        var queryName = "ColumnStructureQuery_" + Guid.NewGuid().ToString("N")[..8];
        var sheetName = "DataSheet";

        // STEP 1: M code for query with ONE column
        string oneColumnMCode = @"let
    Source = #table(
        {""Column1""},
        {
            {""Value1""},
            {""Value2""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        var createResult = _powerQueryCommands.Create(batch, queryName, oneColumnMCode, PowerQueryLoadMode.LoadToTable, sheetName);
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");

        // STEP 2: Verify there is only ONE column
        var rangeCommands = new RangeCommands();
        var usedRange1 = rangeCommands.GetUsedRange(batch, sheetName);
        Assert.True(usedRange1.Success, $"GetUsedRange failed: {usedRange1.ErrorMessage}");
        Assert.Equal(1, usedRange1.ColumnCount);

        // STEP 3: Updated M code (still one column, different data)
        string oneColumnUpdatedMCode = @"let
    Source = #table(
        {""Column1""},
        {
            {""UpdatedValue1""},
            {""UpdatedValue2""},
            {""UpdatedValue3""}
        }
    )
in
    Source";

        // STEP 3: Update query to ONE column (now auto-refreshes)
        var updateResult1 = _powerQueryCommands.Update(batch, queryName, oneColumnUpdatedMCode);
        Assert.True(updateResult1.Success, $"First update failed: {updateResult1.ErrorMessage}");

        // STEP 4: Check that there is still only ONE column
        var usedRange2 = rangeCommands.GetUsedRange(batch, sheetName);
        Assert.True(usedRange2.Success, $"GetUsedRange after first update failed: {usedRange2.ErrorMessage}");
        Assert.Equal(1, usedRange2.ColumnCount);

        // STEP 5: M code for TWO columns
        string twoColumnMCode = @"let
    Source = #table(
        {""Column1"", ""Column2""},
        {
            {""A"", ""B""},
            {""C"", ""D""},
            {""E"", ""F""}
        }
    )
in
    Source";

        // STEP 5: Update query to TWO columns (now auto-refreshes)
        // This validates the fix: PreserveColumnInfo=false allows column structure updates
        var updateResult2 = _powerQueryCommands.Update(batch, queryName, twoColumnMCode);
        Assert.True(updateResult2.Success, $"Second update failed: {updateResult2.ErrorMessage}");

        // STEP 6: Check that there are now TWO columns
        // This validates the fix: PreserveColumnInfo=false allows column structure updates
        var usedRange3 = rangeCommands.GetUsedRange(batch, sheetName);
        Assert.True(usedRange3.Success, $"GetUsedRange after second update failed: {usedRange3.ErrorMessage}");

        // Diagnostic: Capture actual column structure for better error messages
        var values = rangeCommands.GetValues(batch, sheetName, usedRange3.RangeAddress);
        Assert.True(values.Success, $"GetValues failed: {values.ErrorMessage}");

        // Get column headers to see what columns Excel created
        var headerRow = values.Values.FirstOrDefault();
        var columnNames = headerRow != null
            ? string.Join(", ", headerRow.Select(c => c?.ToString() ?? "null"))
            : "No headers found";

        // Primary assertion: Verify column count is correct
        Assert.True(usedRange3.ColumnCount == 2,
            $"Expected 2 columns but got {usedRange3.ColumnCount}. " +
            $"Actual columns: [{columnNames}]");

        // Additional assertion: Verify values match expected structure
        Assert.True(values.ColumnCount == 2,
            $"Expected 2 columns in values but got {values.ColumnCount}. " +
            $"Columns: [{columnNames}]");
    }

    /// <summary>
    /// REGRESSION TEST: Validates SetLoadToTableAsync prevents column accumulation
    ///
    /// Historical bug scenario (delete/recreate workaround):
    /// 1. Create query with 1 column (Column1)
    /// 2. Update M code to 2 columns (Column1, Column2)
    /// 3. Delete query + SetLoadToTable (recreate QueryTable) + Refresh
    /// 4. BUG: Excel created 3 columns (Column1, Column1, Column2) - ACCUMULATION instead of replacing!
    ///
    /// Root cause: Deleting QueryTable left data on worksheet, causing visual concatenation
    /// Fix: Clear worksheet data before creating new QueryTable in SetLoadToTableAsync
    ///
    /// This test reproduces the exact scenario from early testing where we saw accumulated columns.
    /// </summary>
    [Fact]
    public void Update_QueryColumnStructureWithDeleteRecreate_NoAccumulation()
    {
        // Arrange
        var testExcelFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(PowerQueryCommandsTests),
            nameof(Update_QueryColumnStructureWithDeleteRecreate_NoAccumulation),
            _tempDir);

        var queryName = "AccumulationBug_" + Guid.NewGuid().ToString("N")[..8];
        var sheetName = "TestSheet";

        // STEP 1: M code for query with 1 column
        string oneColumnMCode = @"let
    Source = #table(
        {""Column1""},
        {
            {""A""},
            {""B""}
        }
    )
in
    Source";

        using var batch = ExcelSession.BeginBatch(testExcelFile);

        // Import and load to worksheet
        var createResult = _powerQueryCommands.Create(batch, queryName, oneColumnMCode, PowerQueryLoadMode.LoadToTable, sheetName);
        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");

        // Verify initial state: 1 column
        var rangeCommands = new RangeCommands();
        var usedRange1 = rangeCommands.GetUsedRange(batch, sheetName);
        Assert.True(usedRange1.Success);
        Assert.Equal(1, usedRange1.ColumnCount);

        // STEP 2: M code for query with 2 columns
        string twoColumnMCode = @"let
    Source = #table(
        {""Column1"", ""Column2""},
        {
            {""X"", ""Y""},
            {""Z"", ""W""}
        }
    )
in
    Source";

        // STEP 2: Update query to TWO columns (now auto-refreshes)
        var updateResult = _powerQueryCommands.Update(batch, queryName, twoColumnMCode);
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // STEP 3: Apply the DELETE + RECREATE workflow (historically caused 3-column bug)
        var deleteResult = _powerQueryCommands.Delete(batch, queryName);
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");

        var recreateResult = _powerQueryCommands.Create(batch, queryName, twoColumnMCode, PowerQueryLoadMode.ConnectionOnly);
        Assert.True(recreateResult.Success, $"Re-create failed: {recreateResult.ErrorMessage}");

        var loadResult = _powerQueryCommands.LoadTo(batch, queryName, PowerQueryLoadMode.LoadToTable, sheetName, "A1");
        Assert.True(loadResult.Success, $"LoadTo failed: {loadResult.ErrorMessage}");

        var refreshResult = _powerQueryCommands.Refresh(batch, queryName, TimeSpan.FromMinutes(5));
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");

        // STEP 4: Verify NO column accumulation (fix validation)
        var usedRange2 = rangeCommands.GetUsedRange(batch, sheetName);
        Assert.True(usedRange2.Success);

        // Get actual column headers for diagnostic output
        var values = rangeCommands.GetValues(batch, sheetName, usedRange2.RangeAddress);
        Assert.True(values.Success);

        var headerRow = values.Values.FirstOrDefault();
        var columnNames = headerRow != null
            ? string.Join(", ", headerRow.Select(c => c?.ToString() ?? "null"))
            : "No headers found";

        // PRIMARY ASSERTION: Validates the fix prevents column accumulation
        // Should be 2 columns (Column1, Column2), NOT 3 columns (Column1, Column1, Column2)
        Assert.True(usedRange2.ColumnCount == 2,
            $"COLUMN ACCUMULATION DETECTED! Expected 2 columns but got {usedRange2.ColumnCount}. " +
            $"Actual columns: [{columnNames}]. " +
            $"The fix (clearing worksheet before creating QueryTable) should prevent accumulation.");
    }

    #endregion

    #region Helper Methods

    #endregion
}




