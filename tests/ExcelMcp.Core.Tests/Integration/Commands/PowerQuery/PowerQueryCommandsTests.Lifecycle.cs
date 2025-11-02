using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Tests for PowerQuery lifecycle operations: List, Import, View, Export, Update, Delete
/// Uses shared PowerQuery file from fixture for READ operations.
/// WRITE operations use unique query names to avoid conflicts.
/// </summary>
public partial class PowerQueryCommandsTests
{
    /// <summary>
    /// Verifies that listing queries returns the fixture queries.
    /// </summary>
    [Fact]
    public async Task List_FixtureWorkbook_ReturnsFixtureQueries()
    {
        // Act - Use shared file from fixture
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Queries);
        Assert.Equal(3, result.Queries.Count); // Fixture creates 3 queries
    }

    /// <summary>
    /// Verifies that importing a Power Query from a valid M code file succeeds.
    /// Tests the basic import functionality without loading data to worksheet.
    /// </summary>
    [Fact]
    public async Task Import_ValidMCode_ReturnsSuccess()
    {
        // Arrange - Use unique query name
        var queryName = $"Test_{nameof(Import_ValidMCode_ReturnsSuccess)}_{Guid.NewGuid():N}";
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Import_ValidMCode_ReturnsSuccess));

        // Act - Use shared file
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile);
        
        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }

    /// <summary>
    /// Verifies that listing queries after import shows the newly imported query.
    /// Tests the integration between import and list operations.
    /// </summary>
    [Fact]
    public async Task List_AfterImport_IncludesNewQuery()
    {
        // Arrange - Use unique query name
        var queryName = $"Test_{nameof(List_AfterImport_IncludesNewQuery)}_{Guid.NewGuid():N}";
        var testQueryFile = CreateUniqueTestQueryFile(nameof(List_AfterImport_IncludesNewQuery));

        // Act - Use single batch on shared file
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile);
        var result = await _powerQueryCommands.ListAsync(batch);
        
        // Assert
        Assert.True(result.Success);
        Assert.NotNull(result.Queries);
        Assert.Contains(result.Queries, q => q.Name == queryName);
    }

    /// <summary>
    /// Verifies that viewing an existing query returns its M code.
    /// Tests that the query's formula is accessible and contains expected content.
    /// </summary>
    [Fact]
    public async Task View_BasicQuery_ReturnsMCode()
    {
        // Act - View fixture query on shared file
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ViewAsync(batch, "BasicQuery");
        
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
    public async Task Export_BasicQuery_CreatesFile()
    {
        // Arrange
        var exportPath = Path.Join(_tempDir, $"exported_{Guid.NewGuid():N}.pq");

        // Act - Export fixture query from shared file
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        var result = await _powerQueryCommands.ExportAsync(batch, "BasicQuery", exportPath);
        
        // Assert
        Assert.True(result.Success);
        Assert.True(System.IO.File.Exists(exportPath));
    }

    /// <summary>
    /// Verifies that updating an existing query with new M code succeeds.
    /// Tests the update functionality with a simple M code replacement.
    /// </summary>
    [Fact]
    public async Task Update_ExistingQuery_ReturnsSuccess()
    {
        // Arrange - Create unique query to update
        var queryName = $"Test_{nameof(Update_ExistingQuery_ReturnsSuccess)}_{Guid.NewGuid():N}";
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Update_ExistingQuery_ReturnsSuccess));
        var updateFile = Path.Join(_tempDir, $"updated_{Guid.NewGuid():N}.pq");
        System.IO.File.WriteAllText(updateFile, "let\n    UpdatedSource = 1\nin\n    UpdatedSource");

        // Act - Import then update on shared file
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile);
        var result = await _powerQueryCommands.UpdateAsync(batch, queryName, updateFile);
        
        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// Verifies that deleting an existing query succeeds.
    /// Tests the delete operation on a previously imported query.
    /// </summary>
    [Fact]
    public async Task Delete_ExistingQuery_ReturnsSuccess()
    {
        // Arrange - Create unique query to delete
        var queryName = $"Test_{nameof(Delete_ExistingQuery_ReturnsSuccess)}_{Guid.NewGuid():N}";
        var testQueryFile = CreateUniqueTestQueryFile(nameof(Delete_ExistingQuery_ReturnsSuccess));

        // Act - Import then delete on shared file
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile);
        var result = await _powerQueryCommands.DeleteAsync(batch, queryName);
        
        // Assert
        Assert.True(result.Success);
    }

    /// <summary>
    /// Verifies the complete lifecycle: import, delete, then list shows query is gone.
    /// Tests that deletion properly removes the query from the workbook.
    /// </summary>
    [Fact]
    public async Task ImportThenDelete_UniqueQuery_RemovedFromList()
    {
        // Arrange - Create unique query
        var queryName = $"Test_{nameof(ImportThenDelete_UniqueQuery_RemovedFromList)}_{Guid.NewGuid():N}";
        var testQueryFile = CreateUniqueTestQueryFile(nameof(ImportThenDelete_UniqueQuery_RemovedFromList));

        // Act - All operations in single batch on shared file
        await using var batch = await ExcelSession.BeginBatchAsync(_powerQueryFile);
        await _powerQueryCommands.ImportAsync(batch, queryName, testQueryFile);
        await _powerQueryCommands.DeleteAsync(batch, queryName);
        var result = await _powerQueryCommands.ListAsync(batch);

        // Assert - Query should not be in list
        Assert.True(result.Success);
        Assert.DoesNotContain(result.Queries, q => q.Name == queryName);
    }
}
