using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Power Query Core operations.
/// These tests require Excel installation and validate Core Power Query data operations.
/// Tests use Core commands directly (not through CLI wrapper).
///
/// For comprehensive workflow tests (mode switching), see PowerQueryLoadConfigWorkflowTests.cs.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
public class PowerQueryCommandsTests : IDisposable
{
    private readonly IPowerQueryCommands _powerQueryCommands;
    private readonly IFileCommands _fileCommands;
    private readonly ISheetCommands _sheetCommands;
    private readonly string _tempDir;
    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the test class.
    /// Creates a temporary directory for test files.
    /// Each test method creates its own unique Excel file to avoid parallel execution conflicts.
    /// </summary>
    public PowerQueryCommandsTests()
    {
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
        _fileCommands = new FileCommands();
        _sheetCommands = new SheetCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_PQ_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Creates a unique Excel file for a test to avoid parallel execution conflicts.
    /// Each test gets its own isolated file.
    /// </summary>
    private string CreateUniqueTestExcelFile()
    {
        var uniqueFile = Path.Combine(_tempDir, $"TestWorkbook_{Guid.NewGuid():N}.xlsx");
        var result = _fileCommands.CreateEmptyAsync(uniqueFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }
        return uniqueFile;
    }

    /// <summary>
    /// Creates a unique test Power Query M code file.
    /// Each test gets its own isolated M code file.
    /// </summary>
    private string CreateUniqueTestQueryFile()
    {
        var uniqueFile = Path.Combine(_tempDir, $"TestQuery_{Guid.NewGuid():N}.pq");
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

        File.WriteAllText(uniqueFile, mCode);
        return uniqueFile;
    }

    /// <summary>
    /// Verifies that listing queries in a new Excel file returns success with an empty query list.
    /// </summary>
    [Fact]
    public async Task List_WithValidFile_ReturnsSuccessResult()
    {
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile();

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var result = await _powerQueryCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);
        Assert.Empty(result.Queries); // New file has no queries
    }

    /// <summary>
    /// Verifies that importing a Power Query from a valid M code file succeeds.
    /// Tests the basic import functionality without loading data to worksheet.
    /// </summary>
    [Fact]
    public async Task Import_WithValidMCode_ReturnsSuccessResult()
    {
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var result = await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    /// <summary>
    /// Verifies that listing queries after import shows the newly imported query.
    /// Tests the integration between import and list operations.
    /// </summary>
    [Fact]
    public async Task List_AfterImport_ShowsNewQuery()
    {
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act - Use single batch for both operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile);
        var result = await _powerQueryCommands.ListAsync(batch);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Queries);
        Assert.Single(result.Queries);
        Assert.Equal("TestQuery", result.Queries[0].Name);
    }

    /// <summary>
    /// Verifies that viewing an existing query returns its M code.
    /// Tests that the query's formula is accessible and contains expected content.
    /// </summary>
    [Fact]
    public async Task View_WithExistingQuery_ReturnsMCode()
    {
        // Arrange - Use same file for both operations
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act - Use single batch for both operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile);
        var result = await _powerQueryCommands.ViewAsync(batch, "TestQuery");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.MCode);
        Assert.Contains("Source", result.MCode);
    }

    /// <summary>
    /// Verifies that exporting an existing query creates a file with the M code.
    /// Tests that the exported file exists and can be read.
    /// </summary>
    [Fact]
    public async Task Export_WithExistingQuery_CreatesFile()
    {
        // Arrange - Create unique test file ONCE, reuse for all operations
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();
        var exportPath = Path.Combine(_tempDir, "exported.pq");

        // Act - Use single batch for both operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile);
        var result = await _powerQueryCommands.ExportAsync(batch, "TestQuery", exportPath);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(exportPath));
    }

    /// <summary>
    /// Verifies that updating an existing query with new M code succeeds.
    /// Tests the update functionality with a simple M code replacement.
    /// </summary>
    [Fact]
    public async Task Update_WithValidMCode_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file ONCE, reuse for all operations
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();
        var updateFile = Path.Combine(_tempDir, "updated.pq");
        File.WriteAllText(updateFile, "let\n    UpdatedSource = 1\nin\n    UpdatedSource");

        // Act - Use single batch for both operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile);
        var result = await _powerQueryCommands.UpdateAsync(batch, "TestQuery", updateFile);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// Verifies that deleting an existing query succeeds.
    /// Tests the delete operation on a previously imported query.
    /// </summary>
    [Fact]
    public async Task Delete_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file ONCE, reuse for all operations
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act - Use single batch for both operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile);
        var result = await _powerQueryCommands.DeleteAsync(batch, "TestQuery");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// Verifies the complete lifecycle: import, delete, then list shows no queries.
    /// Tests that deletion properly removes the query from the workbook.
    /// </summary>
    [Fact]
    public async Task Import_ThenDelete_ThenList_ShowsEmpty()
    {
        // Arrange - Use single unique file for entire lifecycle test
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act - Import
        await using (var batch = await ExcelSession.BeginBatchAsync(testExcelFile))
        {
            await _powerQueryCommands.ImportAsync(batch, "TestQuery", testQueryFile);
            await batch.SaveAsync();
        }

        // Act - Delete
        await using (var batch = await ExcelSession.BeginBatchAsync(testExcelFile))
        {
            await _powerQueryCommands.DeleteAsync(batch, "TestQuery");
            await batch.SaveAsync();
        }

        // Act - List
        await using (var batch = await ExcelSession.BeginBatchAsync(testExcelFile))
        {
            var result = await _powerQueryCommands.ListAsync(batch);

            // Assert
            Assert.True(result.Success);
            Assert.Empty(result.Queries);
        }
    }

    /// <summary>
    /// Verifies that setting a query to connection-only mode succeeds.
    /// Tests that the query can be configured to not load data anywhere.
    /// </summary>
    [Fact]
    public async Task SetConnectionOnly_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file ONCE
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act - Use single batch for both operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestConnectionOnly", testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var result = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "TestConnectionOnly");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"SetConnectionOnly failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-connection-only", result.Action);
    }

    /// <summary>
    /// Verifies that setting a query to load to table mode succeeds.
    /// Tests atomic operation: configuration AND data loading to worksheet.
    /// Validates that the load configuration is correctly set and data is loaded.
    /// </summary>
    [Fact]
    public async Task SetLoadToTable_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file ONCE
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act - Use single batch for all operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToTable", testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var result = await _powerQueryCommands.SetLoadToTableAsync(batch, "TestLoadToTable", "TestSheet");

        // Assert - Verify the operation succeeded
        Assert.True(result.Success, $"SetLoadToTable failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-table", result.Action);

        // Verify the load configuration was actually set
        var configResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestLoadToTable");
        Assert.True(configResult.Success, $"Failed to get load config: {configResult.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, configResult.LoadMode);

        // Verify sheet was created
        var sheetsResult = await _sheetCommands.ListAsync(batch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "TestSheet");

        // Verify table/QueryTable exists on worksheet (actual data loaded)
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("TestSheet");
            dynamic queryTables = sheet.QueryTables;
            Assert.True(queryTables.Count > 0, "Expected at least one QueryTable on the worksheet");
            return 0;
        });

        // Verify query is NOT in Data Model (LoadToTable only)
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(batch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.DoesNotContain(tablesResult.Tables, t => t.Name == "TestLoadToTable");

        await batch.SaveAsync();
    }

    /// <summary>
    /// Verifies that setting a query to load to data model mode succeeds.
    /// Tests atomic operation: configuration AND data loading to Data Model.
    /// Validates that the load configuration is correctly set and data is loaded to PowerPivot.
    /// Note: Data Model is always available in modern Excel.
    /// </summary>
    [Fact]
    public async Task SetLoadToDataModel_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file ONCE
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act - Use single batch for all operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToDataModel", testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var result = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "TestLoadToDataModel");

        // Assert - Verify the operation succeeded
        Assert.True(result.Success, $"SetLoadToDataModel failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-data-model", result.Action);

        // Verify the load configuration was actually set
        var configResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "TestLoadToDataModel");
        Assert.True(configResult.Success, $"Failed to get load config: {configResult.ErrorMessage}");
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, configResult.LoadMode);

        // Verify query is in Data Model
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(batch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.Contains(tablesResult.Tables, t => t.Name == "TestLoadToDataModel");

        // Verify NO QueryTable on any worksheet (LoadToDataModel only, no worksheet table)
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheets = ctx.Book.Worksheets;
            int sheetCount = sheets.Count;
            for (int i = 1; i <= sheetCount; i++)
            {
                dynamic sheet = sheets.Item(i);
                dynamic queryTables = sheet.QueryTables;
                Assert.True(queryTables.Count == 0, $"Expected no QueryTables on sheet '{sheet.Name}' for LoadToDataModel mode");
            }
            return 0;
        });

        await batch.SaveAsync();
    }

    /// <summary>
    /// Verifies that setting a query to load to both table and data model succeeds.
    /// Tests atomic operation: configuration AND data loading to both destinations.
    /// Validates that data is loaded to both worksheet and Data Model.
    /// </summary>
    [Fact]
    public async Task SetLoadToBoth_WithExistingQuery_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file ONCE
        var testExcelFile = CreateUniqueTestExcelFile();
        var testQueryFile = CreateUniqueTestQueryFile();

        // Act - Use single batch for all operations
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);
        var importResult = await _powerQueryCommands.ImportAsync(batch, "TestLoadToBoth", testQueryFile);
        Assert.True(importResult.Success, $"Failed to import query: {importResult.ErrorMessage}");

        var result = await _powerQueryCommands.SetLoadToBothAsync(batch, "TestLoadToBoth", "TestSheet");

        // Assert - Verify the operation succeeded
        Assert.True(result.Success, $"SetLoadToBoth failed: {result.ErrorMessage}");
        Assert.Equal("pq-set-load-to-both", result.Action);

        // Verify sheet was created with a table
        var sheetsResult = await _sheetCommands.ListAsync(batch);
        Assert.True(sheetsResult.Success);
        Assert.Contains(sheetsResult.Worksheets, w => w.Name == "TestSheet");

        // Verify table exists on worksheet (QueryTable from Power Query)
        await batch.Execute<int>((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item("TestSheet");
            dynamic queryTables = sheet.QueryTables;
            Assert.True(queryTables.Count > 0, "Expected at least one QueryTable on the worksheet");
            return 0;
        });

        // Verify query is in Data Model
        var dataModelCommands = new DataModelCommands();
        var tablesResult = await dataModelCommands.ListTablesAsync(batch);
        Assert.True(tablesResult.Success, $"Failed to list Data Model tables: {tablesResult.ErrorMessage}");
        Assert.Contains(tablesResult.Tables, t => t.Name == "TestLoadToBoth");

        await batch.SaveAsync();
    }

    /// <summary>
    /// Verifies error handling when operating on non-existent queries.
    /// Tests that all operations return appropriate "not found" errors.
    /// NOTE: Comprehensive workflow test for load configuration is in PowerQueryLoadConfigWorkflowTests.cs
    /// </summary>
    [Fact]
    public async Task Operations_WithNonExistentQuery_ReturnNotFoundError()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(CreateUniqueTestExcelFile());

        // Act & Assert - Test multiple operations return "not found" error
        var viewResult = await _powerQueryCommands.ViewAsync(batch, "NonExistentQuery");
        Assert.False(viewResult.Success);
        Assert.Contains("not found", viewResult.ErrorMessage);

        var getConfigResult = await _powerQueryCommands.GetLoadConfigAsync(batch, "NonExistentQuery");
        Assert.False(getConfigResult.Success);
        Assert.Contains("not found", getConfigResult.ErrorMessage);

        var setTableResult = await _powerQueryCommands.SetLoadToTableAsync(batch, "NonExistentQuery", "TestSheet");
        Assert.False(setTableResult.Success);
        Assert.Contains("not found", setTableResult.ErrorMessage);

        var setModelResult = await _powerQueryCommands.SetLoadToDataModelAsync(batch, "NonExistentQuery");
        Assert.False(setModelResult.Success);
        Assert.Contains("not found", setModelResult.ErrorMessage);

        var setBothResult = await _powerQueryCommands.SetLoadToBothAsync(batch, "NonExistentQuery", "TestSheet");
        Assert.False(setBothResult.Success);
        Assert.Contains("not found", setBothResult.ErrorMessage);

        var setConnResult = await _powerQueryCommands.SetConnectionOnlyAsync(batch, "NonExistentQuery");
        Assert.False(setConnResult.Success);
        Assert.Contains("not found", setConnResult.ErrorMessage);
    }

    /// <summary>
    /// Verifies that a Power Query can successfully reference and load data from another Power Query.
    /// Tests the common pattern where one query (DerivedQuery) references another query (SourceQuery).
    /// This validates that:
    /// 1. The source query loads data successfully
    /// 2. The derived query can reference the source query
    /// 3. The derived query loads the expected data from the source
    /// 4. Both queries are independently accessible and functional
    /// </summary>
    [Fact]
    public async Task Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully()
    {
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile();

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

        var sourceQueryFile = Path.Combine(_tempDir, $"SourceQuery_{Guid.NewGuid():N}.pq");
        File.WriteAllText(sourceQueryFile, sourceQueryMCode);

        // Create M code for the derived query (references the source query)
        // This query filters products with price > 15
        string derivedQueryMCode = @"let
    Source = SourceQuery,
    FilteredRows = Table.SelectRows(Source, each [Price] > 15)
in
    FilteredRows";

        var derivedQueryFile = Path.Combine(_tempDir, $"DerivedQuery_{Guid.NewGuid():N}.pq");
        File.WriteAllText(derivedQueryFile, derivedQueryMCode);

        // Act & Assert
        await using var batch = await ExcelSession.BeginBatchAsync(testExcelFile);

        // Step 1: Import the source query (must be imported first)
        var sourceImportResult = await _powerQueryCommands.ImportAsync(
            batch,
            "SourceQuery",
            sourceQueryFile,
            loadDestination: "worksheet");

        Assert.True(sourceImportResult.Success,
            $"Source query import failed: {sourceImportResult.ErrorMessage}");

        // Step 2: Import the derived query (references SourceQuery)
        var derivedImportResult = await _powerQueryCommands.ImportAsync(
            batch,
            "DerivedQuery",
            derivedQueryFile,
            loadDestination: "worksheet");

        Assert.True(derivedImportResult.Success,
            $"Derived query import failed: {derivedImportResult.ErrorMessage}");

        // Step 3: Verify both queries exist in the workbook
        var listResult = await _powerQueryCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Equal(2, listResult.Queries.Count);
        Assert.Contains(listResult.Queries, q => q.Name == "SourceQuery");
        Assert.Contains(listResult.Queries, q => q.Name == "DerivedQuery");

        // Step 4: Verify the source query M code
        var sourceViewResult = await _powerQueryCommands.ViewAsync(batch, "SourceQuery");
        Assert.True(sourceViewResult.Success);
        Assert.Equal("SourceQuery", sourceViewResult.QueryName);
        Assert.Contains("#table", sourceViewResult.MCode);
        Assert.Contains("ProductID", sourceViewResult.MCode);

        // Step 5: Verify the derived query M code references SourceQuery
        var derivedViewResult = await _powerQueryCommands.ViewAsync(batch, "DerivedQuery");
        Assert.True(derivedViewResult.Success);
        Assert.Equal("DerivedQuery", derivedViewResult.QueryName);
        Assert.Contains("SourceQuery", derivedViewResult.MCode);
        Assert.Contains("Table.SelectRows", derivedViewResult.MCode);
        Assert.Contains("Price", derivedViewResult.MCode);

        // Step 6: Refresh both queries to ensure they execute successfully
        var sourceRefreshResult = await _powerQueryCommands.RefreshAsync(batch, "SourceQuery");
        Assert.True(sourceRefreshResult.Success,
            $"Source query refresh failed: {sourceRefreshResult.ErrorMessage}");

        var derivedRefreshResult = await _powerQueryCommands.RefreshAsync(batch, "DerivedQuery");
        Assert.True(derivedRefreshResult.Success,
            $"Derived query refresh failed: {derivedRefreshResult.ErrorMessage}");

        await batch.SaveAsync();
    }

    /// <summary>
    /// Disposes test resources and cleans up temporary directory.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Protected implementation of Dispose pattern.
    /// Cleans up temporary test directory if it exists.
    /// </summary>
    /// <param name="disposing">True if called from Dispose(), false if called from finalizer</param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed) return;

        if (disposing)
        {
            try
            {
                if (Directory.Exists(_tempDir))
                {
                    Directory.Delete(_tempDir, true);
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }

        _disposed = true;
    }
}

