using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model relationship operations (LLM-relevant workflows only)
/// Uses shared Data Model file from fixture.
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListRelationships_FixtureModel_ReturnsRelationships()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(result.Success,
            $"ListRelationships MUST succeed - Data Model is always available. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Relationships);
        Assert.Equal(2, result.Relationships.Count); // Fixture creates exactly 2 relationships
    }

    [Fact]
    public async Task ListRelationships_WithRealisticDataModel_ReturnsRelationshipsWithTables()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert - Fixture creates exactly 2 relationships
        Assert.True(result.Success,
            $"ListRelationships MUST succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Relationships);
        Assert.Equal(2, result.Relationships.Count);

        // Validate SalesTable->CustomersTable relationship
        var salesCustomersRel = result.Relationships.FirstOrDefault(r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable");

        Assert.NotNull(salesCustomersRel);
        Assert.Equal("CustomerID", salesCustomersRel.FromColumn);
        Assert.Equal("CustomerID", salesCustomersRel.ToColumn);
        Assert.True(salesCustomersRel.IsActive, "SalesTable->CustomersTable relationship should be active");

        // Validate SalesTable->ProductsTable relationship
        var salesProductsRel = result.Relationships.FirstOrDefault(r =>
            r.FromTable == "SalesTable" && r.ToTable == "ProductsTable");

        Assert.NotNull(salesProductsRel);
        Assert.Equal("ProductID", salesProductsRel.FromColumn);
        Assert.Equal("ProductID", salesProductsRel.ToColumn);
        Assert.True(salesProductsRel.IsActive, "SalesTable->ProductsTable relationship should be active");
    }

    [Fact]
    public async Task CreateRelationship_ValidTablesAndColumns_CreatesSuccessfully()
    {
        // Arrange - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);

        // First delete existing SalesTable->CustomersTable relationship to allow creating it fresh
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        if (listResult.Success && listResult.Relationships?.Any(r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID") == true)
        {
            // Delete existing relationship to allow creating it fresh
            await _dataModelCommands.DeleteRelationshipAsync(batch, "SalesTable", "CustomerID", "CustomersTable", "CustomerID");
        }

        // Act - Create the relationship
        var createResult = await _dataModelCommands.CreateRelationshipAsync(
            batch,
            "SalesTable",
            "CustomerID",
            "CustomersTable",
            "CustomerID"
        );
        
        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(createResult.Success,
            $"CreateRelationship MUST succeed. Error: {createResult.ErrorMessage}");

        // Verify relationship was created
        var verifyResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        Assert.True(verifyResult.Success);
        Assert.Contains(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");
    }

    [Fact]
    public async Task DeleteRelationship_ExistingRelationship_ReturnsSuccess()
    {
        // Arrange - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);

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
        
        // Assert - MUST succeed (Data Model is always available in Excel 2013+)
        Assert.True(deleteResult.Success,
            $"DeleteRelationship MUST succeed. Error: {deleteResult.ErrorMessage}");

        // Verify relationship was deleted
        var verifyResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        Assert.True(verifyResult.Success);
        Assert.DoesNotContain(verifyResult.Relationships, r =>
            r.FromTable == "SalesTable" && r.ToTable == "CustomersTable" &&
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");
        
        // Recreate it for other tests (since we share the file)
        await _dataModelCommands.CreateRelationshipAsync(batch,
            "SalesTable", "CustomerID", "CustomersTable", "CustomerID", active: true);
    }
}
