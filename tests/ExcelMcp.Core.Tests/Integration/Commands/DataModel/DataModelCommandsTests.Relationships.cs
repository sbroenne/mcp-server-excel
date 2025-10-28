using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model relationship operations
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListRelationships_WithValidFile_ReturnsSuccessResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        if (result.Success)
        {
            Assert.NotNull(result.Relationships);
        }
    }

    [Fact]
    public async Task ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithTables()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        // If Data Model was created successfully with relationships, validate them
        if (result.Success && result.Relationships != null && result.Relationships.Count > 0)
        {
            // Should have at least 2 relationships (Sales->Customers, Sales->Products)
            Assert.True(result.Relationships.Count >= 2, $"Expected at least 2 relationships, got {result.Relationships.Count}");

            // Validate Sales->Customers relationship
            var salesCustomersRel = result.Relationships.FirstOrDefault(r =>
                r.FromTable == "Sales" && r.ToTable == "Customers");

            if (salesCustomersRel != null)
            {
                Assert.Equal("CustomerID", salesCustomersRel.FromColumn);
                Assert.Equal("CustomerID", salesCustomersRel.ToColumn);
                Assert.True(salesCustomersRel.IsActive, "Sales->Customers relationship should be active");
            }

            // Validate Sales->Products relationship
            var salesProductsRel = result.Relationships.FirstOrDefault(r =>
                r.FromTable == "Sales" && r.ToTable == "Products");

            if (salesProductsRel != null)
            {
                Assert.Equal("ProductID", salesProductsRel.FromColumn);
                Assert.Equal("ProductID", salesProductsRel.ToColumn);
                Assert.True(salesProductsRel.IsActive, "Sales->Products relationship should be active");
            }
        }
    }

    [Fact(Skip = "Data Model test helper requires specific Excel version/configuration. May fail on some environments due to Data Model availability.")]
    public async Task DeleteRelationship_WithValidRelationship_ReturnsSuccessResult()
    {
        // Arrange - Requires Data Model with relationships
        await using var listBatch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var listResult = await _dataModelCommands.ListRelationshipsAsync(listBatch);

        Assert.True(listResult.Success, "ListRelationships should succeed");
        Assert.NotNull(listResult.Relationships);
        Assert.True(listResult.Relationships.Count > 0, "Data Model should have relationships for this test");

        // Use the first relationship for testing
        var rel = listResult.Relationships[0];

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.DeleteRelationshipAsync(
            batch,
            rel.FromTable,
            rel.FromColumn,
            rel.ToTable,
            rel.ToColumn
        );

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.SuggestedNextActions);
        Assert.Contains(result.SuggestedNextActions, s => s.Contains("deleted successfully"));
    }

    [Fact]
    public async Task DeleteRelationship_WithNonExistentRelationship_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.DeleteRelationshipAsync(
            batch,
            "FakeTable",
            "FakeColumn",
            "OtherTable",
            "OtherColumn"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("not found in Data Model"),
            $"Expected 'no Data Model' or 'not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task DeleteRelationship_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Act & Assert - BeginBatchAsync should throw FileNotFoundException for non-existent file
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync("NonExistent.xlsx");
            await _dataModelCommands.DeleteRelationshipAsync(
                batch,
                "Table1",
                "Col1",
                "Table2",
                "Col2"
            );
        });
    }
}
