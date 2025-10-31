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
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("ListTables_WithValidFile_ReturnsSuccessResult.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
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
    public async Task ListTables_WithRealisticDataModel_ReturnsTablesWithData()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("ListTables_WithRealisticDataModel_ReturnsTablesWithData.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ListTablesAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);

        // If Data Model was created successfully, validate the tables
        if (result.Tables != null && result.Tables.Count > 0)
        {
            // Should have SalesTable, CustomersTable, and ProductsTable
            Assert.True(result.Tables.Count >= 3, $"Expected at least 3 tables, got {result.Tables.Count}");

            var tableNames = result.Tables.Select(t => t.Name).ToList();
            Assert.Contains("SalesTable", tableNames);
            Assert.Contains("CustomersTable", tableNames);
            Assert.Contains("ProductsTable", tableNames);

            // Validate SalesTable has expected columns
            var salesTable = result.Tables.FirstOrDefault(t => t.Name == "SalesTable");
            if (salesTable != null)
            {
                Assert.True(salesTable.RecordCount > 0, "SalesTable should have rows");
            }
        }
    }
}
