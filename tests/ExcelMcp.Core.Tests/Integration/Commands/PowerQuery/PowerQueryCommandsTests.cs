using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
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
public class PowerQueryCommandsTests : IClassFixture<PowerQueryTestsFixture>
{
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly IFileCommands _fileCommands;
    private readonly ISheetCommands _sheetCommands;
    private readonly string _powerQueryFile;
    private readonly PowerQueryCreationResult _creationResult;
    private readonly string _tempDir;

    public PowerQueryCommandsTests(PowerQueryTestsFixture fixture)
    {
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
        _fileCommands = new FileCommands();
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
    public async Task Import_ValidMCode_ReturnsSuccess()
    {
        // Arrange - Use unique file to avoid polluting fixture
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Import_ValidMCode_ReturnsSuccess),
            _tempDir);
        var queryName = "TestQuery";
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Import_ValidMCode_ReturnsSuccess));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var result = await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile, "connection-only");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    /// <summary>
    /// Tests listing queries in a workbook.
    /// LLM use case: "show me all Power Queries in this workbook"
    /// </summary>
    [Fact]
    public async Task List_FixtureWorkbook_ReturnsFixtureQueries()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ListAsync(batch);

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
    public async Task View_BasicQuery_ReturnsMCode()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ViewAsync(batch, "BasicQuery");

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
    public async Task Update_ExistingQuery_ReturnsSuccess()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Update_ExistingQuery_ReturnsSuccess),
            _tempDir);

        var queryName = "PQ_Update_" + Guid.NewGuid().ToString("N").Substring(0, 8);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Update_ExistingQuery_ReturnsSuccess));
        var updateFile = Path.Join(_tempDir, $"updated_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(updateFile, "let\n    UpdatedSource = 1\nin\n    UpdatedSource");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile);
        var result = await _powerQueryCommands.UpdateAsync(batch, queryName, updateFile);

        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// Tests deleting an existing query.
    /// LLM use case: "delete this Power Query"
    /// </summary>
    [Fact]
    public async Task Delete_ExistingQuery_ReturnsSuccess()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Delete_ExistingQuery_ReturnsSuccess),
            _tempDir);

        var queryName = "PQ_Delete_" + Guid.NewGuid().ToString("N").Substring(0, 8);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Delete_ExistingQuery_ReturnsSuccess));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile);
        var result = await _powerQueryCommands.DeleteAsync(batch, queryName);

        // Assert
        Assert.True(result.Success);
    }

    #endregion

    #region Load Destination Workflows (3 tests)

    /// <summary>
    /// Tests converting connection-only query to worksheet load mode.
    /// LLM use case: "load this query data to a worksheet"
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded),
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded));

        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile, "connection-only");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        // Act - Apply load configuration and refresh
        var setLoadResult = await _powerQueryCommands.SetLoadToTableAsync(batch, "TestQuery", "TestQuery");
        Assert.True(setLoadResult.Success, $"SetLoadToTable failed: {setLoadResult.ErrorMessage}");

        var refreshResult = await _powerQueryCommands.RefreshAsync(batch, "TestQuery");

        // Assert
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        Assert.False(refreshResult.IsConnectionOnly, "Query should no longer be connection-only");

        // Verify worksheet was created with data
        var sheetsResult = await _sheetCommands.ListAsync(batch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "TestQuery");

        // Verify QueryTable exists on worksheet
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("TestQuery");
            dynamic queryTables = sheet.QueryTables;
            Assert.True(queryTables.Count > 0, "Expected at least one QueryTable on the worksheet");
            return 0;
        });
    }

    /// <summary>
    /// Tests converting connection-only query to data model load mode.
    /// LLM use case: "load this query to the data model for DAX"
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationDataModel_LoadsToDataModel()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Refresh_WithLoadDestinationDataModel_LoadsToDataModel),
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_WithLoadDestinationDataModel_LoadsToDataModel));

        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestDMQuery", testQueryFile, "connection-only");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        // Act - Apply load configuration and refresh
        var setLoadResult = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "TestDMQuery");
        Assert.True(setLoadResult.Success, $"SetLoadToDataModel failed: {setLoadResult.ErrorMessage}");

        var refreshResult = await _powerQueryCommands.RefreshAsync(batch, "TestDMQuery");

        // Assert
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        Assert.False(refreshResult.IsConnectionOnly, "Query should no longer be connection-only");

        // Verify query is in Data Model
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(batch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.Contains(tablesResult.Tables, t => t.Name == "TestDMQuery");
    }

    /// <summary>
    /// Tests loading query to both worksheet and data model.
    /// LLM use case: "load this query to both worksheet and data model"
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationBoth_LoadsToBothDestinations()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Refresh_WithLoadDestinationBoth_LoadsToBothDestinations),
            _tempDir);
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Refresh_WithLoadDestinationBoth_LoadsToBothDestinations));

        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestBothQuery", testQueryFile, "connection-only");
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        // Act - Apply load configuration and refresh
        var setLoadResult = await _powerQueryCommands.SetLoadToBothAsync(batch, "TestBothQuery", "TestBothQuery");
        Assert.True(setLoadResult.Success, $"SetLoadToBoth failed: {setLoadResult.ErrorMessage}");

        var refreshResult = await _powerQueryCommands.RefreshAsync(batch, "TestBothQuery");

        // Assert
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");
        Assert.False(refreshResult.IsConnectionOnly, "Query should no longer be connection-only");

        // Verify worksheet was created with data
        var sheetsResult = await _sheetCommands.ListAsync(batch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "TestBothQuery");

        // Verify QueryTable exists on worksheet
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("TestBothQuery");
            dynamic queryTables = sheet.QueryTables;
            Assert.True(queryTables.Count > 0, "Expected at least one QueryTable on the worksheet");
            return 0;
        });

        // Verify query is in Data Model
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(batch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.Contains(tablesResult.Tables, t => t.Name == "TestBothQuery");
    }

    #endregion

    #region Error Handling (1 test)

    /// <summary>
    /// Tests that operations on non-existent queries return proper error messages.
    /// LLM use case: Proper error feedback when query name is wrong
    /// </summary>
    [Fact]
    public async Task Operations_WithNonExistentQuery_ReturnNotFoundError()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Operations_WithNonExistentQuery_ReturnNotFoundError),
            _tempDir);
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);

        // Act & Assert - Multiple operations should return "not found" error
        var viewResult = await _powerQueryCommands.ViewAsync(batch, "NonExistentQuery");
        Assert.False(viewResult.Success);
        Assert.Contains("not found", viewResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var getConfigResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "NonExistentQuery");
        Assert.False(getConfigResult.Success);
        Assert.Contains("not found", getConfigResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var setTableResult = await _powerQueryCommands.SetLoadToTableAsync(batch, "NonExistentQuery", "TestSheet");
        Assert.False(setTableResult.Success);
        Assert.Contains("not found", setTableResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var setModelResult = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "NonExistentQuery");
        Assert.False(setModelResult.Success);
        Assert.Contains("not found", setModelResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var setBothResult = await _powerQueryCommands.SetLoadToBothAsync(batch, "NonExistentQuery", "TestSheet");
        Assert.False(setBothResult.Success);
        Assert.Contains("not found", setBothResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);

        var setConnResult = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "NonExistentQuery");
        Assert.False(setConnResult.Success);
        Assert.Contains("not found", setConnResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Advanced Use Cases (1 test)

    /// <summary>
    /// Tests that one Power Query can successfully reference another Power Query.
    /// LLM use case: "create a query that filters data from another query"
    /// </summary>
    [Fact]
    public async Task Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
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

        var sourceQueryFile = Path.Join(_tempDir, $"SourceQuery_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(sourceQueryFile, sourceQueryMCode);

        // Create M code for the derived query (references the source query)
        string derivedQueryMCode = @"let
    Source = SourceQuery,
    FilteredRows = Table.SelectRows(Source, each [Price] > 15)
in
    FilteredRows";

        var derivedQueryFile = Path.Join(_tempDir, $"DerivedQuery_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(derivedQueryFile, derivedQueryMCode);

        // Act & Assert
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);

        // Import source query first
        var sourceImportResult = await _powerQueryCommands.ImportAsync(
            batch,
            "SourceQuery",
            sourceQueryFile,
            loadDestination: "worksheet");

        Assert.True(sourceImportResult.Success,
            $"Source query import failed: {sourceImportResult.ErrorMessage}");

        // Import derived query (references SourceQuery)
        var derivedImportResult = await _powerQueryCommands.ImportAsync(
            batch,
            "DerivedQuery",
            derivedQueryFile,
            loadDestination: "worksheet");

        Assert.True(derivedImportResult.Success,
            $"Derived query import failed: {derivedImportResult.ErrorMessage}");

        // Verify both queries exist in the workbook
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Equal(2, listResult.Queries.Count);
        Assert.Contains(listResult.Queries, q => q.Name == "SourceQuery");
        Assert.Contains(listResult.Queries, q => q.Name == "DerivedQuery");

        // Verify the derived query M code references SourceQuery
        var derivedViewResult = await _powerQueryCommands.ViewAsync(batch, "DerivedQuery");
        Assert.True(derivedViewResult.Success);
        Assert.Contains("SourceQuery", derivedViewResult.MCode);
        Assert.Contains("Table.SelectRows", derivedViewResult.MCode);

        // Refresh both queries to ensure they execute successfully
        var sourceRefreshResult = await _powerQueryCommands.RefreshAsync(batch, "SourceQuery");
        Assert.True(sourceRefreshResult.Success,
            $"Source query refresh failed: {sourceRefreshResult.ErrorMessage}");

        var derivedRefreshResult = await _powerQueryCommands.RefreshAsync(batch, "DerivedQuery");
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
    /// This test verifies that UpdateAsync preserves the load configuration.
    /// </summary>
    [Fact]
    public async Task Update_QueryLoadedToSheet_PreservesLoadConfiguration()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Update_QueryLoadedToSheet_PreservesLoadConfiguration),
            _tempDir);

        var queryName = "LoadedQuery_" + Guid.NewGuid().ToString("N").Substring(0, 8);
        var sheetName = "DataSheet";
        var initialQueryFile = CreateUniqueTestQueryFile("Initial");
        var updatedQueryFile = Path.Join(_tempDir, $"updated_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(updatedQueryFile,
            @"let
    UpdatedSource = #table(
        {""NewCol1"", ""NewCol2""},
        {
            {""Updated1"", ""Updated2""},
            {""Data1"", ""Data2""}
        }
    )
in
    UpdatedSource");

        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);

        // STEP 1: Import query and load to worksheet
        var importResult = await _powerQueryCommands.ImportAsync(batch, queryName, initialQueryFile, "worksheet", sheetName);
        Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

        // Verify initial load configuration
        var loadConfigBefore = await _powerQueryCommands.GetLoadConfigAsync(batch, queryName);
        Assert.True(loadConfigBefore.Success, "GetLoadConfig before update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadConfigBefore.LoadMode);
        Assert.Equal(sheetName, loadConfigBefore.TargetSheet);

        // STEP 2: Update the query M code
        var updateResult = await _powerQueryCommands.UpdateAsync(batch, queryName, updatedQueryFile);
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // STEP 3: Verify load configuration is PRESERVED (regression check)
        var loadConfigAfter = await _powerQueryCommands.GetLoadConfigAsync(batch, queryName);
        Assert.True(loadConfigAfter.Success, "GetLoadConfig after update failed");

        // THE BUG: This assertion should pass but might fail if UpdateAsync doesn't restore load config
        Assert.Equal(PowerQueryLoadMode.LoadToTable, loadConfigAfter.LoadMode);
        Assert.Equal(sheetName, loadConfigAfter.TargetSheet);

        // STEP 4: Verify data is still loaded to the sheet (not connection-only)
        var listResult = await _sheetCommands.ListAsync(batch);
        Assert.Contains(listResult.Worksheets, s => s.Name == sheetName);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Creates a unique test Power Query M code file.
    /// Used by tests that need to create new queries.
    /// </summary>
    private string CreateUniqueTestQueryFile(string testName)
    {
        var uniqueFile = Path.Join(_tempDir, $"{testName}_{Guid.NewGuid():N}.pq");
        string mCode = @"let
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

        System.IO.File.WriteAllText(uniqueFile, mCode);
        return uniqueFile;
    }

    #endregion
}
