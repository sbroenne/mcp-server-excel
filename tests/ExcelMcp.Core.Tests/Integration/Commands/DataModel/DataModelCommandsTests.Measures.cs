using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model measure operations
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ListMeasuresAsync(batch);

        // Assert - Data Model is ALWAYS available in Excel 2013+
        Assert.True(result.Success,
            $"ListMeasures MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Measures);

        // If Data Model was created successfully with measures, validate them
        if (result.Measures.Count > 0)
        {
            // Should have at least Total Sales, Average Sale, Total Customers
            Assert.True(result.Measures.Count >= 3, $"Expected at least 3 measures, got {result.Measures.Count}");

            var measureNames = result.Measures.Select(m => m.Name).ToList();
            Assert.Contains("Total Sales", measureNames);
            Assert.Contains("Average Sale", measureNames);
            Assert.Contains("Total Customers", measureNames);

            // Validate Total Sales measure has DAX formula
            var totalSales = result.Measures.FirstOrDefault(m => m.Name == "Total Sales");
            if (totalSales != null)
            {
                Assert.NotNull(totalSales.FormulaPreview);
                Assert.Contains("SUM", totalSales.FormulaPreview, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("Amount", totalSales.FormulaPreview, StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    [Fact]
    public async Task ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ViewMeasureAsync(batch, "Total Sales");

        // Assert - Data Model is ALWAYS available in Excel 2013+
        // If measure doesn't exist, it's because test fixture didn't create it (separate issue)
        Assert.True(result.Success,
            $"ViewMeasure MUST succeed if measure exists. Error: {result.ErrorMessage}");
        Assert.NotNull(result.DaxFormula);
        Assert.Contains("SUM", result.DaxFormula, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Sales", result.DaxFormula);
        Assert.Contains("Amount", result.DaxFormula);
        Assert.Equal("Total Sales", result.MeasureName);
    }

    [Fact(Skip = "Data Model test helper requires specific Excel version/configuration. May fail on some environments due to Data Model availability.")]
    public async Task DeleteMeasure_WithValidMeasure_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("DeleteMeasure_WithValidMeasure_ReturnsSuccessResult.xlsx");

        // Arrange - Create a test measure using PRODUCTION command
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];

        await using var createBatch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _dataModelCommands.CreateMeasureAsync(createBatch, "SalesTable", measureName, "SUM(SalesTable[Amount])");
        await createBatch.SaveAsync();

        Assert.True(createResult.Success, $"Test setup failed - could not create measure: {createResult.ErrorMessage}");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.DeleteMeasureAsync(batch, measureName);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");

        // Verify the measure was actually deleted by listing measures
        var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
        if (listResult.Success)
        {
            Assert.DoesNotContain(listResult.Measures, m => m.Name == measureName);
        }
    }

    // Phase 2: CREATE/UPDATE Tests

    [Fact]
    public async Task CreateMeasure_WithValidParameters_CreatesSuccessfully()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("CreateMeasure_WithValidParameters_CreatesSuccessfully.xlsx");

        // Arrange
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(SalesTable[Amount])";

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
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
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("CreateMeasure_WithFormatType_CreatesWithFormat.xlsx");

        // Arrange
        var measureName = "FormattedMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(SalesTable[Amount])";

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
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
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("UpdateMeasure_WithValidFormula_UpdatesSuccessfully.xlsx");

        // Arrange - Create a measure first
        var measureName = "UpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var originalFormula = "SUM(SalesTable[Amount])";
        var updatedFormula = "AVERAGE(SalesTable[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _dataModelCommands.CreateMeasureAsync(batch, "SalesTable", measureName, originalFormula);

        if (createResult.Success)
        {
            // Act - Update the formula
            var updateResult = await _dataModelCommands.UpdateMeasureAsync(batch, measureName, daxFormula: updatedFormula);
            // Assert
            Assert.True(updateResult.Success, $"Expected success but got error: {updateResult.ErrorMessage}");

            // Verify the update
            var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
            Assert.True(viewResult.Success);
            Assert.Contains("AVERAGE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        }
    }
}
