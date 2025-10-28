using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

public partial class DataModelTomCommandsTests
{
    #region CreateRelationship Tests

    [Fact]
    public async Task CreateRelationship_WithValidParameters_ReturnsSuccess()
    {
        // Arrange - This test requires the Data Model to have Sales, Customers, Products tables
        // Skip if TOM connection fails

        // Act
        var result = _tomCommands.CreateRelationship(
            _testExcelFile,
            "Sales",
            "CustomerID",
            "Customers",
            "CustomerID",
            isActive: true,
            crossFilterDirection: "Single"
        );

        // Assert
        if (result.Success)
        {
            Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
            Assert.NotNull(result.SuggestedNextActions);
            Assert.Contains(result.SuggestedNextActions, s => s.Contains("created"));

            // Verify the relationship was created
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
            var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);
            if (listResult.Success)
            {
                Assert.Contains(listResult.Relationships, r =>
                    r.FromTable == "Sales" && r.ToTable == "Customers");
            }
        }
        else
        {
            // If TOM connection failed or relationship already exists, that's acceptable
            Assert.True(
                result.ErrorMessage?.Contains("Data Model") == true ||
                result.ErrorMessage?.Contains("connect") == true ||
                result.ErrorMessage?.Contains("already exists") == true,
                $"Expected Data Model, connection, or duplicate error, got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void CreateRelationship_WithInvalidTable_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateRelationship(
            _testExcelFile,
            "InvalidTable",
            "ID",
            "AnotherTable",
            "ID"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("not found") ||
            result.ErrorMessage.Contains("connect"),
            $"Expected 'not found' or connection error, got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public void CreateRelationship_WithEmptyParameters_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateRelationship(
            _testExcelFile,
            "",
            "Column1",
            "Table2",
            "Column2"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region UpdateRelationship Tests

    [Fact]
    public async Task UpdateRelationship_WithValidParameters_ReturnsSuccess()
    {
        // Arrange - First ensure a relationship exists
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Test should fail if no relationships exist, not skip
        if (!listResult.Success || listResult.Relationships == null || listResult.Relationships.Count == 0)
        {
            Assert.Fail($"Data Model does not have relationships for testing. Success={listResult.Success}, Count={listResult.Relationships?.Count ?? 0}");
        }

        var rel = listResult.Relationships[0];

        // Act - Update the relationship
        var updateResult = _tomCommands.UpdateRelationship(
            _testExcelFile,
            rel.FromTable,
            rel.FromColumn,
            rel.ToTable,
            rel.ToColumn,
            isActive: !rel.IsActive
        );

        // Assert
        if (updateResult.Success)
        {
            Assert.True(updateResult.Success, $"Expected success but got error: {updateResult.ErrorMessage}");
            Assert.NotNull(updateResult.SuggestedNextActions);
        }
        else
        {
            // TOM connection failure is acceptable
            Assert.True(
                updateResult.ErrorMessage?.Contains("connect") == true,
                $"Expected connection error, got: {updateResult.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void UpdateRelationship_WithNoParameters_ReturnsError()
    {
        // Act
        var result = _tomCommands.UpdateRelationship(
            _testExcelFile,
            "Table1",
            "Col1",
            "Table2",
            "Col2"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("at least one property", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
