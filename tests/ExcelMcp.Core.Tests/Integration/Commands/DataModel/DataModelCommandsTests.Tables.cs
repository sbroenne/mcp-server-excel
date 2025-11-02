using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model table operations
/// Uses shared Data Model file from fixture.
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListTables_WithValidFile_ReturnsSuccessResult()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ListTablesAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        Assert.Equal(3, result.Tables.Count); // Fixture creates exactly 3 tables
    }

    [Fact]
    public async Task ListTables_WithRealisticDataModel_ReturnsTablesWithData()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ListTablesAsync(batch);

        // Assert - Fixture creates exactly 3 tables
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        Assert.Equal(3, result.Tables.Count);

        var tableNames = result.Tables.Select(t => t.Name).ToList();
        Assert.Contains("SalesTable", tableNames);
        Assert.Contains("CustomersTable", tableNames);
        Assert.Contains("ProductsTable", tableNames);

        // Validate SalesTable has expected columns
        var salesTable = result.Tables.FirstOrDefault(t => t.Name == "SalesTable");
        Assert.NotNull(salesTable);
        Assert.True(salesTable.RecordCount > 0, "SalesTable should have rows");
    }
}
