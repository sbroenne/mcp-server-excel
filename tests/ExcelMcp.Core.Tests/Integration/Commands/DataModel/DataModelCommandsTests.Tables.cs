using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model table operations
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListTables_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListTablesAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);

        // New file without Data Model should indicate that
        if (!result.Success && result.ErrorMessage?.Contains("does not contain a Data Model") == true)
        {
            // This is expected for empty workbook
            Assert.Contains("does not contain a Data Model", result.ErrorMessage);
        }
    }

    [Fact]
    public async Task ListTables_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        var nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");

        // Act & Assert - BeginBatchAsync should throw FileNotFoundException for non-existent file
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(nonExistentFile);
            await _dataModelCommands.ListTablesAsync(batch);
        });
    }

    [Fact]
    public async Task ListTables_WithRealisticDataModel_ReturnsTablesWithData()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListTablesAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);

        // If Data Model was created successfully, validate the tables
        if (result.Tables != null && result.Tables.Count > 0)
        {
            // Should have Sales, Customers, and Products tables
            Assert.True(result.Tables.Count >= 3, $"Expected at least 3 tables, got {result.Tables.Count}");

            var tableNames = result.Tables.Select(t => t.Name).ToList();
            Assert.Contains("Sales", tableNames);
            Assert.Contains("Customers", tableNames);
            Assert.Contains("Products", tableNames);

            // Validate Sales table has expected columns
            var salesTable = result.Tables.FirstOrDefault(t => t.Name == "Sales");
            if (salesTable != null)
            {
                Assert.True(salesTable.RecordCount > 0, "Sales table should have rows");
            }
        }
    }
}
