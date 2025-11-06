using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.ComInterop;
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

    /// <summary>
    /// LLM use case: "update a query to change column structure and verify columns update"
    ///
    /// Scenario:
    /// 1. Create a PowerQuery with one column and load to a worksheet
    /// 2. Check that there is only one column
    /// 3. Update the query and load again
    /// 4. Check that there is only one column
    /// 5. Update the query to create two columns and load again
    /// 6. Check that there are two columns
    /// </summary>
    [Fact]
    public async Task Update_QueryColumnStructure_UpdatesWorksheetColumns()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Update_QueryColumnStructure_UpdatesWorksheetColumns),
            _tempDir);

        var queryName = "ColumnStructureQuery_" + Guid.NewGuid().ToString("N").Substring(0, 8);
        var sheetName = "DataSheet";

        // STEP 1: Create query with ONE column
        var oneColumnQueryFile = Path.Join(_tempDir, $"onecolumn_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(oneColumnQueryFile,
            @"let
    Source = #table(
        {""Column1""},
        {
            {""Value1""},
            {""Value2""}
        }
    )
in
    Source");

        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);

        var importResult = await _powerQueryCommands.ImportAsync(batch, queryName, oneColumnQueryFile, "worksheet", sheetName);
        Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

        // STEP 2: Verify there is only ONE column
        var rangeCommands = new RangeCommands();
        var usedRange1 = await rangeCommands.GetUsedRangeAsync(batch, sheetName);
        Assert.True(usedRange1.Success, $"GetUsedRange failed: {usedRange1.ErrorMessage}");
        Assert.Equal(1, usedRange1.ColumnCount);

        // STEP 3: Update query (still one column, different data)
        var oneColumnUpdatedFile = Path.Join(_tempDir, $"onecolumn_updated_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(oneColumnUpdatedFile,
            @"let
    Source = #table(
        {""Column1""},
        {
            {""UpdatedValue1""},
            {""UpdatedValue2""},
            {""UpdatedValue3""}
        }
    )
in
    Source");

        var updateResult1 = await _powerQueryCommands.UpdateAsync(batch, queryName, oneColumnUpdatedFile);
        Assert.True(updateResult1.Success, $"First update failed: {updateResult1.ErrorMessage}");

        // Refresh to reload data
        var refreshResult1 = await _powerQueryCommands.RefreshAsync(batch, queryName);
        Assert.True(refreshResult1.Success, $"First refresh failed: {refreshResult1.ErrorMessage}");

        // STEP 4: Verify QueryTable still exists (not converted to range) and check column count
        var queryTableCheck1 = await batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? currentSheet = null;
                    try
                    {
                        currentSheet = sheets.Item(i);
                        if (currentSheet.Name == sheetName)
                        {
                            sheet = currentSheet;
                            currentSheet = null;
                            break;
                        }
                    }
                    finally
                    {
                        if (currentSheet != null)
                            ComUtilities.Release(ref currentSheet);
                    }
                }

                if (sheet == null)
                    return (false, 0);

                queryTables = sheet.QueryTables;
                int qtCount = queryTables.Count;
                return (qtCount > 0, qtCount);
            }
            finally
            {
                if (queryTables != null)
                    ComUtilities.Release(ref queryTables);
                if (sheet != null)
                    ComUtilities.Release(ref sheet);
                if (sheets != null)
                    ComUtilities.Release(ref sheets);
            }
        });

        Assert.True(queryTableCheck1.Item1, "After first update, QueryTable should still exist (not converted to range)");
        Assert.True(queryTableCheck1.Item2 == 1,
            $"After first update, expected exactly 1 QueryTable but found {queryTableCheck1.Item2}. " +
            "Multiple QueryTables indicates improper cleanup during UpdateAsync.");

        // Check that there is still only ONE column
        var usedRange2 = await rangeCommands.GetUsedRangeAsync(batch, sheetName);
        Assert.True(usedRange2.Success, $"GetUsedRange after first update failed: {usedRange2.ErrorMessage}");
        Assert.Equal(1, usedRange2.ColumnCount);

        // STEP 5: Update the query to create TWO columns
        var twoColumnQueryFile = Path.Join(_tempDir, $"twocolumn_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(twoColumnQueryFile,
            @"let
    Source = #table(
        {""Column1"", ""Column2""},
        {
            {""A"", ""B""},
            {""C"", ""D""},
            {""E"", ""F""}
        }
    )
in
    Source");

        var updateResult2 = await _powerQueryCommands.UpdateAsync(batch, queryName, twoColumnQueryFile);
        Assert.True(updateResult2.Success, $"Second update failed: {updateResult2.ErrorMessage}");

        // Refresh to reload data with new column structure
        var refreshResult2 = await _powerQueryCommands.RefreshAsync(batch, queryName);
        Assert.True(refreshResult2.Success, $"Second refresh failed: {refreshResult2.ErrorMessage}");

        // STEP 6: Verify QueryTable still exists (not converted to range) and check column structure
        var queryTableCheck2 = await batch.Execute((ctx, ct) =>
        {
            dynamic? sheets = null;
            dynamic? sheet = null;
            dynamic? queryTables = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? currentSheet = null;
                    try
                    {
                        currentSheet = sheets.Item(i);
                        if (currentSheet.Name == sheetName)
                        {
                            sheet = currentSheet;
                            currentSheet = null;
                            break;
                        }
                    }
                    finally
                    {
                        if (currentSheet != null)
                            ComUtilities.Release(ref currentSheet);
                    }
                }

                if (sheet == null)
                    return (false, 0);

                queryTables = sheet.QueryTables;
                int qtCount = queryTables.Count;
                return (qtCount > 0, qtCount);
            }
            finally
            {
                if (queryTables != null)
                    ComUtilities.Release(ref queryTables);
                if (sheet != null)
                    ComUtilities.Release(ref sheet);
                if (sheets != null)
                    ComUtilities.Release(ref sheets);
            }
        });

        Assert.True(queryTableCheck2.Item1, "After second update, QueryTable should still exist (not converted to range)");
        Assert.True(queryTableCheck2.Item2 == 1,
            $"After second update, expected exactly 1 QueryTable but found {queryTableCheck2.Item2}. " +
            "Multiple QueryTables indicates improper cleanup during UpdateAsync.");

        // Check that there are now TWO columns
        // BUG: This assertion will FAIL because Excel's QueryTable doesn't update column structure on refresh
        var usedRange3 = await rangeCommands.GetUsedRangeAsync(batch, sheetName);
        Assert.True(usedRange3.Success, $"GetUsedRange after second update failed: {usedRange3.ErrorMessage}");

        // Diagnostic: Capture actual column structure before assertion
        var values = await rangeCommands.GetValuesAsync(batch, sheetName, usedRange3.RangeAddress);
        Assert.True(values.Success, $"GetValues failed: {values.ErrorMessage}");

        // Get column headers to see what columns Excel created
        var headerRow = values.Values.FirstOrDefault();
        var columnNames = headerRow != null
            ? string.Join(", ", headerRow.Select(c => c?.ToString() ?? "null"))
            : "No headers found";

        // This will fail and show us what the actual columns are
        Assert.True(usedRange3.ColumnCount == 2,
            $"Expected 2 columns but got {usedRange3.ColumnCount}. " +
            $"Actual columns: [{columnNames}]");

        // Additional assertion that will also fail
        Assert.True(values.ColumnCount == 2,
            $"Expected 2 columns in values but got {values.ColumnCount}. " +
            $"Columns: [{columnNames}]");
    }

    /// <summary>
    /// REGRESSION TEST: Verifies column accumulation bug is FIXED when using delete/recreate workaround
    ///
    /// Original bug scenario:
    /// 1. Create query with 1 column (Column1)
    /// 2. Update M code to 2 columns (Column1, Column2)
    /// 3. Delete query + SetLoadToTable (recreate QueryTable) + Refresh
    /// 4. BUG (FIXED): Excel created 3-4 columns (Column1, Column1, Column2) - ACCUMULATED instead of replacing!
    ///
    /// Root cause: SetLoadToTableAsync called usedRange.Clear() before creating QueryTable
    /// Fix: Removed usedRange.Clear(), only delete query-specific QueryTables
    ///
    /// This test verifies the bug is NOW FIXED and columns are NOT accumulated.
    /// </summary>
    [Fact]
    public async Task Update_QueryColumnStructureWithDeleteRecreate_AccumulatesColumns()
    {
        // Arrange
        var testExcelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(PowerQueryCommandsTests),
            nameof(Update_QueryColumnStructureWithDeleteRecreate_AccumulatesColumns),
            _tempDir);

        var queryName = "AccumulationBug_" + Guid.NewGuid().ToString("N").Substring(0, 8);
        var sheetName = "TestSheet";

        // STEP 1: Create query with 1 column
        var oneColumnFile = Path.Join(_tempDir, $"initial_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(oneColumnFile,
            @"let
    Source = #table(
        {""Column1""},
        {
            {""A""},
            {""B""}
        }
    )
in
    Source");

        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);

        // Import and load to worksheet
        var importResult = await _powerQueryCommands.ImportAsync(batch, queryName, oneColumnFile, "worksheet", sheetName);
        Assert.True(importResult.Success, $"Import failed: {importResult.ErrorMessage}");

        // Verify initial state: 1 column
        var rangeCommands = new RangeCommands();
        var usedRange1 = await rangeCommands.GetUsedRangeAsync(batch, sheetName);
        Assert.True(usedRange1.Success);
        Assert.Equal(1, usedRange1.ColumnCount);

        // STEP 2: Update M code to 2 columns
        var twoColumnFile = Path.Join(_tempDir, $"updated_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(twoColumnFile,
            @"let
    Source = #table(
        {""Column1"", ""Column2""},
        {
            {""X"", ""Y""},
            {""Z"", ""W""}
        }
    )
in
    Source");

        var updateResult = await _powerQueryCommands.UpdateAsync(batch, queryName, twoColumnFile);
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");

        // STEP 3: Apply the DELETE + RECREATE workaround (this causes the 3-column bug!)
        var deleteResult = await _powerQueryCommands.DeleteAsync(batch, queryName);
        Assert.True(deleteResult.Success, $"Delete failed: {deleteResult.ErrorMessage}");

        var reimportResult = await _powerQueryCommands.ImportAsync(batch, queryName, twoColumnFile, "connection-only");
        Assert.True(reimportResult.Success, $"Re-import failed: {reimportResult.ErrorMessage}");

        var setLoadResult = await _powerQueryCommands.SetLoadToTableAsync(batch, queryName, sheetName);
        Assert.True(setLoadResult.Success, $"SetLoadToTable failed: {setLoadResult.ErrorMessage}");

        var refreshResult = await _powerQueryCommands.RefreshAsync(batch, queryName);
        Assert.True(refreshResult.Success, $"Refresh failed: {refreshResult.ErrorMessage}");

        // STEP 4: Verify the bug is FIXED - should have exactly 2 columns (NOT accumulated)
        var usedRange2 = await rangeCommands.GetUsedRangeAsync(batch, sheetName);
        Assert.True(usedRange2.Success);

        // Get actual column headers for diagnostics
        var values = await rangeCommands.GetValuesAsync(batch, sheetName, usedRange2.RangeAddress);
        Assert.True(values.Success);

        var headerRow = values.Values.FirstOrDefault();
        var columnNames = headerRow != null
            ? string.Join(", ", headerRow.Select(c => c?.ToString() ?? "null"))
            : "No headers found";

        // REGRESSION ASSERTION: Verify bug is FIXED - exactly 2 columns (NOT 3-4 due to accumulation)
        // Original bug: Would get 3-4 columns (Column1, Column1, Column2) or (Column1, Column2, Column1, Column2)
        // Expected after fix: Exactly 2 columns (Column1, Column2)
        Assert.True(usedRange2.ColumnCount == 2,
            $"REGRESSION: Column accumulation bug should be FIXED! Expected 2 columns but got {usedRange2.ColumnCount}. " +
            $"Actual columns: [{columnNames}]. " +
            $"If this fails with >2 columns, the bug has regressed. " +
            $"If this fails with <2 columns, something else is wrong.");
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

    /// <summary>
    /// Creates a test query file with custom M code content.
    /// Returns absolute path to .pq file.
    /// </summary>
    private string CreateTestQueryFileWithContent(string uniqueName, string mCode)
    {
        var fileName = $"{uniqueName}_{Guid.NewGuid():N}.pq";
        var filePath = Path.Combine(_tempDir, fileName);
        System.IO.File.WriteAllText(filePath, mCode);
        return filePath;
    }

    #endregion
}
