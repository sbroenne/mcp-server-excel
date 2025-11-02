using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model measure operations
/// Uses shared Data Model file from fixture.
/// Write operations use unique names to avoid conflicts.
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ListMeasuresAsync(batch);

        // Assert - Fixture creates exactly 3 measures
        Assert.True(result.Success,
            $"ListMeasures MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Measures);
        Assert.Equal(3, result.Measures.Count);

        var measureNames = result.Measures.Select(m => m.Name).ToList();
        Assert.Contains("Total Sales", measureNames);
        Assert.Contains("Average Sale", measureNames);
        Assert.Contains("Total Customers", measureNames);

        // Validate Total Sales measure has DAX formula
        var totalSales = result.Measures.FirstOrDefault(m => m.Name == "Total Sales");
        Assert.NotNull(totalSales);
        Assert.NotNull(totalSales.FormulaPreview);
        Assert.Contains("SUM", totalSales.FormulaPreview, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Amount", totalSales.FormulaPreview, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula()
    {
        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.ViewMeasureAsync(batch, "Total Sales");

        // Assert - Fixture creates "Total Sales" measure
        Assert.True(result.Success,
            $"ViewMeasure MUST succeed if measure exists. Error: {result.ErrorMessage}");
        Assert.NotNull(result.DaxFormula);
        Assert.Contains("SUM", result.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Sales", result.DaxFormula);
        Assert.Contains("Amount", result.DaxFormula);
        Assert.Equal("Total Sales", result.MeasureName);
    }

    [Fact]
    public async Task CreateMeasure_WithValidParameters_CreatesSuccessfully()
    {
        // Arrange - Use unique measure name to avoid conflicts with other tests
        var measureName = $"Test_{nameof(CreateMeasure_WithValidParameters_CreatesSuccessfully)}_{Guid.NewGuid():N}";
        var daxFormula = "SUM(SalesTable[Amount])";

        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, daxFormula);
        
        // Assert - Data Model is ALWAYS available in Excel 2013+
        Assert.True(result.Success,
            $"CreateMeasure MUST succeed with valid parameters. Error: {result.ErrorMessage}");

        // Verify measure was created by listing measures
        var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Measures, m => m.Name == measureName);
    }

    [Fact]
    public async Task CreateMeasure_WithFormatType_CreatesWithFormat()
    {
        // Arrange - Use unique measure name
        var measureName = $"Test_{nameof(CreateMeasure_WithFormatType_CreatesWithFormat)}_{Guid.NewGuid():N}";
        var daxFormula = "SUM(SalesTable[Amount])";

        // Act - Use shared data model file
        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, daxFormula,
                                                                 formatType: "Currency", description: "Test measure with currency format");
        
        // Assert - Data Model is ALWAYS available in Excel 2013+
        Assert.True(result.Success,
            $"CreateMeasure with format MUST succeed. Error: {result.ErrorMessage}");

        // Verify measure exists
        var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
        Assert.True(viewResult.Success);
        Assert.Equal(measureName, viewResult.MeasureName);
    }

    [Fact]
    public async Task UpdateMeasure_WithValidFormula_UpdatesSuccessfully()
    {
        // Arrange - Create a unique measure first
        var measureName = $"Test_{nameof(UpdateMeasure_WithValidFormula_UpdatesSuccessfully)}_{Guid.NewGuid():N}";
        var originalFormula = "SUM(SalesTable[Amount])";
        var updatedFormula = "AVERAGE(SalesTable[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);
        var createResult = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, originalFormula);
        Assert.True(createResult.Success, $"Test setup failed - could not create measure: {createResult.ErrorMessage}");

        // Act - Update the formula
        var updateResult = await _dataModelCommands.UpdateMeasureAsync(batch, measureName, daxFormula: updatedFormula);
        
        // Assert
        Assert.True(updateResult.Success, $"Expected success but got error: {updateResult.ErrorMessage}");

        // Verify the update
        var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
        Assert.True(viewResult.Success);
        Assert.Contains("AVERAGE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteMeasure_WithValidMeasure_ReturnsSuccessResult()
    {
        // Arrange - Create a unique measure to delete
        var measureName = $"Test_{nameof(DeleteMeasure_WithValidMeasure_ReturnsSuccessResult)}_{Guid.NewGuid():N}";

        await using var batch = await ExcelSession.BeginBatchAsync(_dataModelFile);

        // Create the measure
        var createResult = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, "SUM(SalesTable[Amount])");
        Assert.True(createResult.Success, $"Test setup failed - could not create measure: {createResult.ErrorMessage}");
        
        // Act - Delete the measure
        var result = await _dataModelCommands.DeleteMeasureAsync(batch, measureName);
        
        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");

        // Verify the measure was actually deleted
        var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Measures, m => m.Name == measureName);
    }
}
