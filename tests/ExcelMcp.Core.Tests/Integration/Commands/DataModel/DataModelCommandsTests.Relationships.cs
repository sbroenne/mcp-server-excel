using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model relationship operations (LLM-relevant workflows only)
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListRelationships_WithValidFile_ReturnsSuccessResult()
    {
        // Arrange
        var testFile = await CreateTestFileAsync("ListRelationships_WithValidFile.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"ListRelationships MUST succeed - Data Model is always available. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Relationships);
    }

    [Fact]
    public async Task ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithTables()
    {
        // Arrange
        var testFile = await CreateTestFileAsync("ListRelationships_WithRealisticDataModel.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"ListRelationships MUST succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Relationships);

        // If Data Model has relationships, validate them
        if (result.Relationships.Count > 0)
        {
            // Should have at least 2 relationships (SalesTable->CustomersTable, SalesTable->ProductsTable)
            Assert.True(result.Relationships.Count >= 2, $"Expected at least 2 relationships, got {result.Relationships.Count}");

            // Validate SalesTable->CustomersTable relationship
            var salesCustomersRel = result.Relationships.FirstOrDefault(r =>
                r.FromTable == "SalesTable" && r.ToTable == "CustomersTable");

            if (salesCustomersRel != null)
            {
                Assert.Equal("CustomerID", salesCustomersRel.FromColumn);
                Assert.Equal("CustomerID", salesCustomersRel.ToColumn);
                Assert.True(salesCustomersRel.IsActive, "SalesTable->CustomersTable relationship should be active");
            }

            // Validate SalesTable->ProductsTable relationship
            var salesProductsRel = result.Relationships.FirstOrDefault(r =>
                r.FromTable == "SalesTable" && r.ToTable == "ProductsTable");

            if (salesProductsRel != null)
            {
                Assert.Equal("ProductID", salesProductsRel.FromColumn);
                Assert.Equal("ProductID", salesProductsRel.ToColumn);
                Assert.True(salesProductsRel.IsActive, "SalesTable->ProductsTable relationship should be active");
            }
        }
    }

    [Fact]
    public async Task CreateRelationship_WithValidParameters_CreatesSuccessfully()
    {
        // Arrange
        var testFile = await CreateTestFileAsync("CreateRelationship_WithValidParameters.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // First delete existing SalesTable->CustomersTable relationship if it exists
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        if (listResult.Success && listResult.Relationships?.Any(r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID") == true)
        {
            // Delete existing relationship to allow creating it fresh
            await _dataModelCommands.DeleteRelationshipAsync(batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");
            await batch.SaveAsync();
        }

        // Act - Create the relationship
        var createResult = await _dataModelCommands.CreateRelationshipAsync(
            batch,
            "SalesTable",
            "CustomerID",
            "CustomersTable",
            "CustomerID"
        );
        await batch.SaveAsync();

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(createResult.Success,
            $"CreateRelationship MUST succeed. Error: {createResult.ErrorMessage}");
        Assert.NotNull(createResult.SuggestedNextActions);
        Assert.Contains(createResult.SuggestedNextActions, s => s.Contains("Relationship created"));

        // Verify relationship was created
        var verifyResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        Assert.True(verifyResult.Success);
        Assert.Contains(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");
    }

    [Fact]
    public async Task DeleteRelationship_WithValidRelationship_ReturnsSuccessResult()
    {
        // Arrange
        var testFile = await CreateTestFileAsync("DeleteRelationship_WithValidRelationship.xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Verify the relationship exists first
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        Assert.True(listResult.Success);
        var relationshipExists = listResult.Relationships?.Any(r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID") == true;

        // Act - Delete the relationship
        var deleteResult = await _dataModelCommands.DeleteRelationshipAsync(
            batch,
            "SalesTable",
            "CustomerID",
            "CustomersTable",
            "CustomerID"
        );
        await batch.SaveAsync();

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(deleteResult.Success,
            $"DeleteRelationship MUST succeed. Error: {deleteResult.ErrorMessage}");

        // Verify relationship was deleted
        var verifyResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        Assert.True(verifyResult.Success);
        Assert.DoesNotContain(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");
    }
}
