using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Integration tests for ExcelPowerQueryTool refresh with loadDestination parameter
/// Validates the bug fix for refresh action respecting loadDestination
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public class ExcelPowerQueryRefreshTests : IDisposable
{
    private readonly string _tempDir;

    public ExcelPowerQueryRefreshTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"PQ_Refresh_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
        {
            try
            {
                Directory.Delete(_tempDir, recursive: true);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Creates a simple test Power Query M code file
    /// </summary>
    private string CreateTestQueryFile(string testName)
    {
        var queryFile = Path.Combine(_tempDir, $"{testName}_{Guid.NewGuid():N}.pq");
        File.WriteAllText(queryFile, "let\n    Source = {1, 2, 3}\nin\n    Source");
        return queryFile;
    }

    /// <summary>
    /// Regression test for bug: refresh with loadDestination='worksheet' should convert
    /// connection-only query to loaded query
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded()
    {
        // Arrange - Create empty workbook
        var testFile = Path.Combine(_tempDir, $"{nameof(Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded)}.xlsx");
        var createResult = await ExcelFileTool.ExcelFile("create-empty", testFile);
        var createData = JsonSerializer.Deserialize<OperationResult>(createResult);
        Assert.NotNull(createData);
        Assert.True(createData.Success, $"Create file failed: {createData.ErrorMessage}");

        // Import query as connection-only
        var queryFile = CreateTestQueryFile(nameof(Refresh_WithLoadDestinationWorksheet_ConvertsConnectionOnlyToLoaded));
        var importResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "import",
            excelPath: testFile,
            queryName: "TestQuery",
            sourcePath: queryFile,
            loadDestination: "connection-only");

        var importData = JsonSerializer.Deserialize<OperationResult>(importResult);
        Assert.NotNull(importData);
        Assert.True(importData.Success, $"Import failed: {importData.ErrorMessage}");

        // Verify query is connection-only
        var configBeforeResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "get-load-config",
            excelPath: testFile,
            queryName: "TestQuery");

        var configBefore = JsonSerializer.Deserialize<PowerQueryLoadConfigResult>(configBeforeResult);
        Assert.NotNull(configBefore);
        Assert.True(configBefore.Success, "Get config failed");
        Assert.Equal(PowerQueryLoadMode.ConnectionOnly, configBefore.LoadMode);

        // Act - Refresh with loadDestination='worksheet'
        var refreshResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "refresh",
            excelPath: testFile,
            queryName: "TestQuery",
            loadDestination: "worksheet");

        // Assert
        var refreshData = JsonSerializer.Deserialize<PowerQueryRefreshResult>(refreshResult);
        Assert.NotNull(refreshData);
        Assert.True(refreshData.Success, $"Refresh with loadDestination failed: {refreshData.ErrorMessage}");
        
        // BUG FIX VERIFICATION: IsConnectionOnly should now be FALSE
        Assert.False(refreshData.IsConnectionOnly, "Query should NO LONGER be connection-only after refresh with loadDestination='worksheet'");

        // Verify load configuration changed
        var configAfterResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "get-load-config",
            excelPath: testFile,
            queryName: "TestQuery");

        var configAfter = JsonSerializer.Deserialize<PowerQueryLoadConfigResult>(configAfterResult);
        Assert.NotNull(configAfter);
        Assert.True(configAfter.Success, "Get config after refresh failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, configAfter.LoadMode);
    }

    /// <summary>
    /// Verifies refresh with loadDestination='data-model' loads query to Data Model
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationDataModel_LoadsToDataModel()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, $"{nameof(Refresh_WithLoadDestinationDataModel_LoadsToDataModel)}.xlsx");
        var createResult = await ExcelFileTool.ExcelFile("create-empty", testFile);
        var createData = JsonSerializer.Deserialize<OperationResult>(createResult);
        Assert.NotNull(createData);
        Assert.True(createData.Success);

        // Import query as connection-only
        var queryFile = CreateTestQueryFile(nameof(Refresh_WithLoadDestinationDataModel_LoadsToDataModel));
        var importResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "import",
            excelPath: testFile,
            queryName: "DataModelQuery",
            sourcePath: queryFile,
            loadDestination: "connection-only");

        var importData = JsonSerializer.Deserialize<OperationResult>(importResult);
        Assert.NotNull(importData);
        Assert.True(importData.Success);

        // Act - Refresh with loadDestination='data-model'
        var refreshResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "refresh",
            excelPath: testFile,
            queryName: "DataModelQuery",
            loadDestination: "data-model");

        // Assert
        var refreshData = JsonSerializer.Deserialize<PowerQueryRefreshResult>(refreshResult);
        Assert.NotNull(refreshData);
        Assert.True(refreshData.Success, $"Refresh failed: {refreshData.ErrorMessage}");
        Assert.False(refreshData.IsConnectionOnly);

        // Verify load configuration
        var configResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "get-load-config",
            excelPath: testFile,
            queryName: "DataModelQuery");

        var config = JsonSerializer.Deserialize<PowerQueryLoadConfigResult>(configResult);
        Assert.NotNull(config);
        Assert.True(config.Success);
        Assert.Equal(PowerQueryLoadMode.LoadToDataModel, config.LoadMode);
    }

    /// <summary>
    /// Verifies refresh with loadDestination='both' loads to both destinations
    /// </summary>
    [Fact]
    public async Task Refresh_WithLoadDestinationBoth_LoadsToBothDestinations()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, $"{nameof(Refresh_WithLoadDestinationBoth_LoadsToBothDestinations)}.xlsx");
        var createResult = await ExcelFileTool.ExcelFile("create-empty", testFile);
        var createData = JsonSerializer.Deserialize<OperationResult>(createResult);
        Assert.NotNull(createData);
        Assert.True(createData.Success);

        // Import query as connection-only
        var queryFile = CreateTestQueryFile(nameof(Refresh_WithLoadDestinationBoth_LoadsToBothDestinations));
        var importResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "import",
            excelPath: testFile,
            queryName: "BothQuery",
            sourcePath: queryFile,
            loadDestination: "connection-only");

        var importData = JsonSerializer.Deserialize<OperationResult>(importResult);
        Assert.NotNull(importData);
        Assert.True(importData.Success);

        // Act - Refresh with loadDestination='both'
        var refreshResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "refresh",
            excelPath: testFile,
            queryName: "BothQuery",
            loadDestination: "both");

        // Assert
        var refreshData = JsonSerializer.Deserialize<PowerQueryRefreshResult>(refreshResult);
        Assert.NotNull(refreshData);
        Assert.True(refreshData.Success, $"Refresh failed: {refreshData.ErrorMessage}");
        Assert.False(refreshData.IsConnectionOnly);

        // Verify load configuration
        var configResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "get-load-config",
            excelPath: testFile,
            queryName: "BothQuery");

        var config = JsonSerializer.Deserialize<PowerQueryLoadConfigResult>(configResult);
        Assert.NotNull(config);
        Assert.True(config.Success);
        Assert.Equal(PowerQueryLoadMode.LoadToBoth, config.LoadMode);
    }

    /// <summary>
    /// Verifies backwards compatibility: refresh without loadDestination parameter
    /// maintains existing behavior
    /// </summary>
    [Fact]
    public async Task Refresh_WithoutLoadDestination_MaintainsExistingBehavior()
    {
        // Arrange - Create query loaded to worksheet
        var testFile = Path.Combine(_tempDir, $"{nameof(Refresh_WithoutLoadDestination_MaintainsExistingBehavior)}.xlsx");
        var createResult = await ExcelFileTool.ExcelFile("create-empty", testFile);
        var createData = JsonSerializer.Deserialize<OperationResult>(createResult);
        Assert.NotNull(createData);
        Assert.True(createData.Success);

        // Import query with default loadDestination (worksheet)
        var queryFile = CreateTestQueryFile(nameof(Refresh_WithoutLoadDestination_MaintainsExistingBehavior));
        var importResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "import",
            excelPath: testFile,
            queryName: "RefreshTest",
            sourcePath: queryFile,
            loadDestination: "worksheet");

        var importData = JsonSerializer.Deserialize<OperationResult>(importResult);
        Assert.NotNull(importData);
        Assert.True(importData.Success);

        // Act - Refresh WITHOUT loadDestination parameter
        var refreshResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "refresh",
            excelPath: testFile,
            queryName: "RefreshTest",
            loadDestination: null);  // Explicit null to test backwards compatibility

        // Assert - Should refresh successfully
        var refreshData = JsonSerializer.Deserialize<PowerQueryRefreshResult>(refreshResult);
        Assert.NotNull(refreshData);
        Assert.True(refreshData.Success, $"Refresh failed: {refreshData.ErrorMessage}");
        Assert.False(refreshData.IsConnectionOnly);
    }

    /// <summary>
    /// Verifies connection-only query without loadDestination remains connection-only
    /// </summary>
    [Fact]
    public async Task Refresh_ConnectionOnlyWithoutLoadDestination_RemainsConnectionOnly()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, $"{nameof(Refresh_ConnectionOnlyWithoutLoadDestination_RemainsConnectionOnly)}.xlsx");
        var createResult = await ExcelFileTool.ExcelFile("create-empty", testFile);
        var createData = JsonSerializer.Deserialize<OperationResult>(createResult);
        Assert.NotNull(createData);
        Assert.True(createData.Success);

        // Import query as connection-only
        var queryFile = CreateTestQueryFile(nameof(Refresh_ConnectionOnlyWithoutLoadDestination_RemainsConnectionOnly));
        var importResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "import",
            excelPath: testFile,
            queryName: "ConnOnlyQuery",
            sourcePath: queryFile,
            loadDestination: "connection-only");

        var importData = JsonSerializer.Deserialize<OperationResult>(importResult);
        Assert.NotNull(importData);
        Assert.True(importData.Success);

        // Act - Refresh without loadDestination
        var refreshResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "refresh",
            excelPath: testFile,
            queryName: "ConnOnlyQuery",
            loadDestination: null);

        // Assert - Should remain connection-only
        var refreshData = JsonSerializer.Deserialize<PowerQueryRefreshResult>(refreshResult);
        Assert.NotNull(refreshData);
        Assert.True(refreshData.Success);
        Assert.True(refreshData.IsConnectionOnly, "Query should remain connection-only when no loadDestination is specified");
    }

    /// <summary>
    /// Verifies custom targetSheet parameter works with refresh
    /// </summary>
    [Fact]
    public async Task Refresh_WithCustomTargetSheet_CreatesCorrectSheet()
    {
        // Arrange
        var testFile = Path.Combine(_tempDir, $"{nameof(Refresh_WithCustomTargetSheet_CreatesCorrectSheet)}.xlsx");
        var createResult = await ExcelFileTool.ExcelFile("create-empty", testFile);
        var createData = JsonSerializer.Deserialize<OperationResult>(createResult);
        Assert.NotNull(createData);
        Assert.True(createData.Success);

        // Import query as connection-only
        var queryFile = CreateTestQueryFile(nameof(Refresh_WithCustomTargetSheet_CreatesCorrectSheet));
        var importResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "import",
            excelPath: testFile,
            queryName: "CustomSheetTest",
            sourcePath: queryFile,
            loadDestination: "connection-only");

        var importData = JsonSerializer.Deserialize<OperationResult>(importResult);
        Assert.NotNull(importData);
        Assert.True(importData.Success);

        // Act - Refresh with custom targetSheet
        var refreshResult = await ExcelPowerQueryTool.ExcelPowerQuery(
            action: "refresh",
            excelPath: testFile,
            queryName: "CustomSheetTest",
            loadDestination: "worksheet",
            targetSheet: "MyCustomSheet");

        // Assert
        var refreshData = JsonSerializer.Deserialize<PowerQueryRefreshResult>(refreshResult);
        Assert.NotNull(refreshData);
        Assert.True(refreshData.Success, $"Refresh failed: {refreshData.ErrorMessage}");

        // Verify worksheet created with custom name
        var listResult = await ExcelWorksheetTool.ExcelWorksheet(
            action: "list",
            excelPath: testFile);

        var sheetList = JsonSerializer.Deserialize<WorksheetListResult>(listResult);
        Assert.NotNull(sheetList);
        Assert.True(sheetList.Success);
        Assert.Contains(sheetList.Worksheets, w => w.Name == "MyCustomSheet");
    }
}
