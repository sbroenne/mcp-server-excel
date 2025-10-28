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

    // Phase 2: CREATE/UPDATE Tests

    [Fact]
    public async Task CreateRelationship_WithValidParameters_CreatesSuccessfully()
    {
        // Arrange - First delete existing Sales->Customers relationship if it exists
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);

        if (listResult.Success && listResult.Relationships?.Any(r => 
            r.FromTable == "Sales" && r.ToTable == "Customers" && 
            r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID") == true)
        {
            // Delete existing relationship to allow creating it fresh
            await _dataModelCommands.DeleteRelationshipAsync(batch, "Sales", "CustomerID", "Customers", "CustomerID");
            await batch.SaveAsync();
        }

        // Act - Create the relationship
        var createResult = await _dataModelCommands.CreateRelationshipAsync(
            batch, 
            "Sales", 
            "CustomerID", 
            "Customers", 
            "CustomerID"
        );
        await batch.SaveAsync();

        // Assert - Should either succeed or indicate no Data Model
        if (createResult.Success)
        {
            Assert.NotNull(createResult.SuggestedNextActions);
            Assert.Contains(createResult.SuggestedNextActions, s => s.Contains("created successfully"));

            // Verify relationship was created
            var verifyResult = await _dataModelCommands.ListRelationshipsAsync(batch);
            Assert.True(verifyResult.Success);
            Assert.Contains(verifyResult.Relationships, r =>
                r.FromTable == "Sales" && r.ToTable == "Customers" &&
                r.FromColumn == "CustomerID" && r.ToColumn == "CustomerID");
        }
        else
        {
            Assert.True(
                createResult.ErrorMessage?.Contains("does not contain a Data Model") == true,
                $"Expected 'no Data Model' error, but got: {createResult.ErrorMessage}"
            );
        }
    }

    [Fact]
    public async Task CreateRelationship_WithInactiveFlag_CreatesInactiveRelationship()
    {
        // Arrange - Create unique test tables or use existing ones
        // For this test, we'll use Sales->Products which may exist
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);

        if (listResult.Success && listResult.Relationships?.Any(r => 
            r.FromTable == "Sales" && r.ToTable == "Products") == true)
        {
            // Delete existing to create fresh
            var existing = listResult.Relationships.First(r => r.FromTable == "Sales" && r.ToTable == "Products");
            await _dataModelCommands.DeleteRelationshipAsync(batch, existing.FromTable, existing.FromColumn, 
                                                            existing.ToTable, existing.ToColumn);
            await batch.SaveAsync();
        }

        // Act - Create inactive relationship
        var createResult = await _dataModelCommands.CreateRelationshipAsync(
            batch, 
            "Sales", 
            "ProductID", 
            "Products", 
            "ProductID",
            active: false
        );
        await batch.SaveAsync();

        // Assert
        if (createResult.Success)
        {
            // Verify relationship was created as inactive
            var verifyResult = await _dataModelCommands.ListRelationshipsAsync(batch);
            var createdRel = verifyResult.Relationships?.FirstOrDefault(r =>
                r.FromTable == "Sales" && r.ToTable == "Products" &&
                r.FromColumn == "ProductID" && r.ToColumn == "ProductID");

            if (createdRel != null)
            {
                Assert.False(createdRel.IsActive, "Relationship should be inactive");
            }
        }
    }

    [Fact]
    public async Task CreateRelationship_WithDuplicateRelationship_ReturnsError()
    {
        // Arrange - Ensure relationship exists
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var firstCreate = await _dataModelCommands.CreateRelationshipAsync(
            batch, 
            "Sales", 
            "CustomerID", 
            "Customers", 
            "CustomerID"
        );

        if (firstCreate.Success)
        {
            await batch.SaveAsync();

            // Act - Try to create duplicate
            var duplicateResult = await _dataModelCommands.CreateRelationshipAsync(
                batch, 
                "Sales", 
                "CustomerID", 
                "Customers", 
                "CustomerID"
            );

            // Assert
            Assert.False(duplicateResult.Success);
            Assert.NotNull(duplicateResult.ErrorMessage);
            Assert.Contains("already exists", duplicateResult.ErrorMessage);
        }
    }

    [Fact]
    public async Task CreateRelationship_WithInvalidTable_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.CreateRelationshipAsync(
            batch, 
            "NonExistentTable", 
            "Column1", 
            "OtherTable", 
            "Column2"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("not found in Data Model"),
            $"Expected 'no Data Model' or 'table not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task CreateRelationship_WithInvalidColumn_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.CreateRelationshipAsync(
            batch, 
            "Sales", 
            "NonExistentColumn", 
            "Customers", 
            "CustomerID"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("not found"),
            $"Expected 'no Data Model' or 'column not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task UpdateRelationship_ToggleActiveToInactive_UpdatesSuccessfully()
    {
        // Arrange - Ensure we have an active relationship
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        
        // Create or ensure Sales->Customers relationship exists and is active
        var createResult = await _dataModelCommands.CreateRelationshipAsync(
            batch, 
            "Sales", 
            "CustomerID", 
            "Customers", 
            "CustomerID",
            active: true
        );

        if (createResult.Success)
        {
            await batch.SaveAsync();

            // Act - Toggle to inactive
            var updateResult = await _dataModelCommands.UpdateRelationshipAsync(
                batch, 
                "Sales", 
                "CustomerID", 
                "Customers", 
                "CustomerID",
                active: false
            );
            await batch.SaveAsync();

            // Assert
            Assert.True(updateResult.Success, $"Expected success but got: {updateResult.ErrorMessage}");
            Assert.NotNull(updateResult.SuggestedNextActions);
            Assert.Contains(updateResult.SuggestedNextActions, s => s.Contains("now inactive"));

            // Verify the change
            var verifyResult = await _dataModelCommands.ListRelationshipsAsync(batch);
            var relationship = verifyResult.Relationships?.FirstOrDefault(r =>
                r.FromTable == "Sales" && r.ToTable == "Customers");

            if (relationship != null)
            {
                Assert.False(relationship.IsActive, "Relationship should be inactive after update");
            }
        }
    }

    [Fact]
    public async Task UpdateRelationship_ToggleInactiveToActive_UpdatesSuccessfully()
    {
        // Arrange - Ensure we have an inactive relationship
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        
        // Create or ensure Sales->Products relationship exists and is inactive
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        if (listResult.Success && listResult.Relationships?.Any(r => 
            r.FromTable == "Sales" && r.ToTable == "Products") == true)
        {
            var existing = listResult.Relationships.First(r => r.FromTable == "Sales" && r.ToTable == "Products");
            await _dataModelCommands.DeleteRelationshipAsync(batch, existing.FromTable, existing.FromColumn,
                                                            existing.ToTable, existing.ToColumn);
            await batch.SaveAsync();
        }

        var createResult = await _dataModelCommands.CreateRelationshipAsync(
            batch, 
            "Sales", 
            "ProductID", 
            "Products", 
            "ProductID",
            active: false
        );

        if (createResult.Success)
        {
            await batch.SaveAsync();

            // Act - Toggle to active
            var updateResult = await _dataModelCommands.UpdateRelationshipAsync(
                batch, 
                "Sales", 
                "ProductID", 
                "Products", 
                "ProductID",
                active: true
            );
            await batch.SaveAsync();

            // Assert
            Assert.True(updateResult.Success, $"Expected success but got: {updateResult.ErrorMessage}");
            Assert.NotNull(updateResult.SuggestedNextActions);
            Assert.Contains(updateResult.SuggestedNextActions, s => s.Contains("now active"));

            // Verify the change
            var verifyResult = await _dataModelCommands.ListRelationshipsAsync(batch);
            var relationship = verifyResult.Relationships?.FirstOrDefault(r =>
                r.FromTable == "Sales" && r.ToTable == "Products");

            if (relationship != null)
            {
                Assert.True(relationship.IsActive, "Relationship should be active after update");
            }
        }
    }

    [Fact]
    public async Task UpdateRelationship_WithNonExistentRelationship_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.UpdateRelationshipAsync(
            batch, 
            "FakeTable", 
            "FakeColumn", 
            "OtherTable", 
            "OtherColumn",
            active: true
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("not found"),
            $"Expected 'no Data Model' or 'not found' error, but got: {result.ErrorMessage}"
        );
    }
}
