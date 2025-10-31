using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for PowerQuery lifecycle operations: List, Import, View, Export, Update, Delete
/// </summary>
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Verifies that listing queries in a new Excel file returns success with an empty query list.
    /// </summary>
    [Fact]
    public async Task List_WithValidFile_ReturnsSuccessResult()
    {
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile(nameof(List_WithValidFile_ReturnsSuccessResult));

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
        var testExcelFile = CreateUniqueTestExcelFile(nameof(Import_WithValidMCode_ReturnsSuccessResult));
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Import_WithValidMCode_ReturnsSuccessResult));

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
        var testExcelFile = CreateUniqueTestExcelFile(nameof(List_AfterImport_ShowsNewQuery));
        var testQueryFile = CreateUniqueTestQueryFile(nameof(List_AfterImport_ShowsNewQuery));

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
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile(nameof(View_WithExistingQuery_ReturnsMCode));
        var testQueryFile = CreateUniqueTestQueryFile(nameof(View_WithExistingQuery_ReturnsMCode));

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
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile(nameof(Export_WithExistingQuery_CreatesFile));
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Export_WithExistingQuery_CreatesFile));
        var exportPath = Path.Combine(_tempDir, $"exported_{Guid.NewGuid():N}.pq");

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
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile(nameof(Update_WithValidMCode_ReturnsSuccessResult));
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Update_WithValidMCode_ReturnsSuccessResult));
        var updateFile = Path.Combine(_tempDir, $"updated_{Guid.NewGuid():N}.pq");
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
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile(nameof(Delete_WithExistingQuery_ReturnsSuccessResult));
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Delete_WithExistingQuery_ReturnsSuccessResult));

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
        // Arrange
        var testExcelFile = CreateUniqueTestExcelFile(nameof(Import_ThenDelete_ThenList_ShowsEmpty));
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Import_ThenDelete_ThenList_ShowsEmpty));

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

        // Act - List and verify
        await using (var batch = await ExcelSession.BeginBatchAsync(testExcelFile))
        {
            var result = await _powerQueryCommands.ListAsync(batch);

            // Assert
            Assert.True(result.Success);
            Assert.Empty(result.Queries);
        }
    }
}
