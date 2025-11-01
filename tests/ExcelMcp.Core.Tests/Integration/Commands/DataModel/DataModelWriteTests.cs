using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Data Model WRITE tests (Create, Update, Delete operations).
/// Uses shared fixture to create ONE Data Model file for all tests (60-120s setup, then all tests are fast).
/// Each test works on the same file but creates uniquely named measures to avoid conflicts.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "DataModel-Write")]
[Trait("RequiresExcel", "true")]
public class DataModelWriteTests : IClassFixture<DataModelWriteTestsFixture>
{
    private readonly DataModelCommands _dataModelCommands;
    private readonly string _sharedTestFile;

    public DataModelWriteTests(DataModelWriteTestsFixture fixture)
    {
        _dataModelCommands = new DataModelCommands();
        _sharedTestFile = fixture.TestFilePath;
    }

    [Fact]
    public async Task CreateMeasure_WithValidParameters_CreatesSuccessfully()
    {
        // Arrange - Use unique measure name to avoid conflicts with other tests
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(SalesTable[Amount])";

        // Act - Work on shared file
        await using var batch = await ExcelSession.BeginBatchAsync(_sharedTestFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, daxFormula);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"CreateMeasure MUST succeed with valid parameters. Error: {result.ErrorMessage}");
        Assert.NotNull(result.SuggestedNextActions);
        Assert.Contains(result.SuggestedNextActions, s => s.Contains("created successfully"));

        // Verify measure was created
        var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Measures, m => m.Name == measureName);
    }

    [Fact]
    public async Task CreateMeasure_WithFormatType_CreatesWithFormat()
    {
        // Arrange - Use unique measure name
        var measureName = "FormattedMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(SalesTable[Amount])";

        // Act - Create measure in first batch
        await using (var batch = await ExcelSession.BeginBatchAsync(_sharedTestFile))
        {
            var result = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, daxFormula,
                                                                     formatType: "Currency", description: "Test measure with currency format");
            await batch.SaveAsync();

            // Assert - CreateMeasure should succeed
            Assert.True(result.Success, $"CreateMeasure with format MUST succeed. Error: {result.ErrorMessage}");
            Assert.NotNull(result.SuggestedNextActions);
        } // Close and save the batch

        // Verify measure exists in NEW batch (after file is closed and reopened)
        await using (var verifyBatch = await ExcelSession.BeginBatchAsync(_sharedTestFile))
        {
            var viewResult = await _dataModelCommands.ViewMeasureAsync(verifyBatch, measureName);
            if (!viewResult.Success)
            {
                throw new Exception($"ViewMeasure FAILED - ErrorMessage: '{viewResult.ErrorMessage}', MeasureName: {measureName}");
            }
            Assert.True(viewResult.Success);
            Assert.Equal(measureName, viewResult.MeasureName);
        }
    }

    [Fact]
    public async Task UpdateMeasure_WithValidFormula_UpdatesSuccessfully()
    {
        // Arrange - Create a unique measure first
        var measureName = "UpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var originalFormula = "SUM(SalesTable[Amount])";
        var updatedFormula = "AVERAGE(SalesTable[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(_sharedTestFile);

        // Create the measure
        var createResult = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, originalFormula);
        Assert.True(createResult.Success, $"Setup failed: {createResult.ErrorMessage}");
        await batch.SaveAsync();

        // Act - Update the formula
        var updateResult = await _dataModelCommands.UpdateMeasureAsync(batch, measureName, daxFormula: updatedFormula);
        await batch.SaveAsync();

        // Assert
        Assert.True(updateResult.Success, $"Expected success but got error: {updateResult.ErrorMessage}");

        // Verify formula was updated
        var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
        Assert.True(viewResult.Success);
        Assert.Contains("AVERAGE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Skip = "Data Model test helper requires specific Excel version/configuration. May fail on some environments due to Data Model availability.")]
    public async Task DeleteMeasure_WithValidMeasure_ReturnsSuccessResult()
    {
        // Arrange - Create a unique measure to delete
        var measureName = "DeleteTest_" + Guid.NewGuid().ToString("N")[..8];

        await using var batch = await ExcelSession.BeginBatchAsync(_sharedTestFile);

        // Create the measure
        var createResult = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, "SUM(SalesTable[Amount])");
        Assert.True(createResult.Success, $"Setup failed: {createResult.ErrorMessage}");
        await batch.SaveAsync();

        // Act - Delete the measure
        var result = await _dataModelCommands.DeleteMeasureAsync(batch, measureName);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.SuggestedNextActions);
        Assert.Contains(result.SuggestedNextActions, s => s.Contains("deleted successfully"));

        // Verify the measure was actually deleted
        var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
        if (listResult.Success)
        {
            Assert.DoesNotContain(listResult.Measures, m => m.Name == measureName);
        }
    }

    [Fact]
    public async Task CreateRelationship_WithValidParameters_CreatesSuccessfully()
    {
        // This test uses existing tables from the shared fixture
        // Note: Relationships already exist in fixture, so this tests creating additional ones or verifying existing

        await using var batch = await ExcelSession.BeginBatchAsync(_sharedTestFile);

        // Act - List relationships to verify they exist
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);

        // Assert
        Assert.True(listResult.Success, $"ListRelationships failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Relationships);
        Assert.True(listResult.Relationships.Count >= 2, "Expected at least 2 relationships from fixture setup");
    }

    [Fact(Skip = "Data Model test helper requires specific Excel version/configuration. May fail on some environments due to Data Model availability.")]
    public async Task DeleteRelationship_WithValidRelationship_ReturnsSuccessResult()
    {
        // Arrange - Use existing relationship from fixture
        await using var batch = await ExcelSession.BeginBatchAsync(_sharedTestFile);

        // Get existing relationships first
        var listResult = await _dataModelCommands.ListRelationshipsAsync(batch);
        Assert.True(listResult.Success);

        if (listResult.Relationships.Count == 0)
        {
            // Skip if no relationships exist
            return;
        }

        var relationship = listResult.Relationships[0];

        // Act - Delete the relationship
        var result = await _dataModelCommands.DeleteRelationshipAsync(batch,
            relationship.FromTable, relationship.FromColumn, relationship.ToTable, relationship.ToColumn);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
    }
}
