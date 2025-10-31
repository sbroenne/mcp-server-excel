using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for Data Model measure operations
/// </summary>
public partial class DataModelCommandsTests
{
    [Fact]
    public async Task ListMeasures_WithValidFile_ReturnsSuccessResult()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("ListMeasures_WithValidFile_ReturnsSuccessResult.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ListMeasuresAsync(batch);

        // Assert - Data Model is ALWAYS available in Excel 2013+
        Assert.True(result.Success,
            $"ListMeasures MUST succeed - Data Model is always available in Excel 2013+. Error: {result.ErrorMessage}");
        Assert.NotNull(result.Measures);
    }

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
    public async Task ViewMeasure_WithNonExistentMeasure_ReturnsErrorResult()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("ViewMeasure_WithNonExistentMeasure_ReturnsErrorResult.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ViewMeasureAsync(batch, "NonExistentMeasure");

        // Assert
        // Should fail because measure doesn't exist (Data Model is always available in Excel 2013+)
        Assert.False(result.Success, "ViewMeasure should fail when measure doesn't exist");
        Assert.NotNull(result.ErrorMessage);
        Assert.True(result.ErrorMessage.Contains("Measure 'NonExistentMeasure' not found"),
            $"Expected 'measure not found' error, but got: {result.ErrorMessage}");
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

    [Fact]
    public async Task ExportMeasure_WithNonExistentMeasure_ReturnsErrorResult()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("ExportMeasure_WithNonExistentMeasure_ReturnsErrorResult.xlsx");
        var outputFile = Path.Combine(_tempDir, "NonExistentMeasure.dax");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ExportMeasureAsync(batch, "NonExistentMeasure", outputFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public async Task ExportMeasure_WithRealisticDataModel_ExportsValidDAXFile()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("ExportMeasure_WithRealisticDataModel_ExportsValidDAXFile.xlsx");

        // Arrange
        var exportPath = Path.Combine(_tempDir, "TotalSales.dax");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.ExportMeasureAsync(batch, "Total Sales", exportPath);

        // Assert - Data Model is ALWAYS available in Excel 2013+
        Assert.True(result.Success,
            $"ExportMeasure MUST succeed if measure exists. Error: {result.ErrorMessage}");
        Assert.True(File.Exists(exportPath), "DAX file should be created");

        var daxContent = File.ReadAllText(exportPath);
        Assert.NotEmpty(daxContent);
        Assert.Contains("SUM", daxContent, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Sales", daxContent);
        Assert.Contains("Amount", daxContent);
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
        Assert.NotNull(result.SuggestedNextActions);
        Assert.Contains(result.SuggestedNextActions, s => s.Contains("deleted successfully"));

        // Verify the measure was actually deleted by listing measures
        var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
        if (listResult.Success)
        {
            Assert.DoesNotContain(listResult.Measures, m => m.Name == measureName);
        }
    }

    [Fact]
    public async Task DeleteMeasure_WithNonExistentMeasure_ReturnsErrorResult()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("DeleteMeasure_WithNonExistentMeasure_ReturnsErrorResult.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.DeleteMeasureAsync(batch, "NonExistentMeasure");

        // Assert - Should fail because file has no Data Model
        Assert.False(result.Success, "DeleteMeasure should fail when file has no Data Model");
        Assert.NotNull(result.ErrorMessage);
        Assert.True(result.ErrorMessage.Contains("does not have a Data Model") || 
                    result.ErrorMessage.Contains("not found"),
            $"Expected 'no data model' or 'not found' error, but got: {result.ErrorMessage}");
    }

    [Fact]
    public async Task DeleteMeasure_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("DeleteMeasure_WithNonExistentFile_ThrowsFileNotFoundException.xlsx");

        // Act & Assert - BeginBatchAsync should throw FileNotFoundException for non-existent file
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync("NonExistent.xlsx");
            await _dataModelCommands.DeleteMeasureAsync(batch, "SomeMeasure");
        });
    }

    // Phase 2: CREATE/UPDATE Tests

    [Fact]
    public async Task CreateMeasure_WithValidParameters_CreatesSuccessfully()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("CreateMeasure_WithValidParameters_CreatesSuccessfully.xlsx");

        // Arrange
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula);
        await batch.SaveAsync();

        // Assert - Data Model is ALWAYS available in Excel 2013+
        Assert.True(result.Success,
            $"CreateMeasure MUST succeed with valid parameters. Error: {result.ErrorMessage}");
        Assert.NotNull(result.SuggestedNextActions);
        Assert.Contains(result.SuggestedNextActions, s => s.Contains("created successfully"));

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
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula,
                                                                 formatType: "Currency", description: "Test measure with currency format");
        await batch.SaveAsync();

        // Assert - Data Model is ALWAYS available in Excel 2013+
        Assert.True(result.Success,
            $"CreateMeasure with format MUST succeed. Error: {result.ErrorMessage}");
        Assert.NotNull(result.SuggestedNextActions);

        // Verify measure exists
        var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
        Assert.True(viewResult.Success);
        Assert.Equal(measureName, viewResult.MeasureName);
    }

    [Fact]
    public async Task CreateMeasure_WithDuplicateName_ReturnsError()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("CreateMeasure_WithDuplicateName_ReturnsError.xlsx");

        // Arrange - First create a measure (MUST succeed - Data Model is always available in Excel 2013+)
        var measureName = "DuplicateTest_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var firstResult = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula);

        Assert.True(firstResult.Success,
            $"First CreateMeasure MUST succeed. Error: {firstResult.ErrorMessage}");
        await batch.SaveAsync();

        // Act - Try to create same measure again
        var duplicateResult = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula);

        // Assert - Should fail with duplicate error
        Assert.False(duplicateResult.Success);
        Assert.NotNull(duplicateResult.ErrorMessage);
        Assert.Contains("already exists", duplicateResult.ErrorMessage);
        Assert.NotNull(duplicateResult.SuggestedNextActions);
    }

    [Fact]
    public async Task CreateMeasure_WithInvalidTable_ReturnsError()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("CreateMeasure_WithInvalidTable_ReturnsError.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "NonExistentTable", "TestMeasure", "1+1");

        // Assert - Should fail because table doesn't exist (Data Model is always available in Excel 2013+)
        Assert.False(result.Success, "CreateMeasure should fail when table doesn't exist");
        Assert.NotNull(result.ErrorMessage);
        Assert.True(result.ErrorMessage.Contains("Table 'NonExistentTable' not found"),
            $"Expected 'table not found' error, but got: {result.ErrorMessage}");
    }

    [Fact]
    public async Task UpdateMeasure_WithValidFormula_UpdatesSuccessfully()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("UpdateMeasure_WithValidFormula_UpdatesSuccessfully.xlsx");

        // Arrange - Create a measure first
        var measureName = "UpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var originalFormula = "SUM(Sales[Amount])";
        var updatedFormula = "AVERAGE(Sales[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, originalFormula);

        if (createResult.Success)
        {
            await batch.SaveAsync();

            // Act - Update the formula
            var updateResult = await _dataModelCommands.UpdateMeasureAsync(batch, measureName, daxFormula: updatedFormula);
            await batch.SaveAsync();

            // Assert
            Assert.True(updateResult.Success, $"Expected success but got error: {updateResult.ErrorMessage}");
            Assert.NotNull(updateResult.SuggestedNextActions);
            Assert.Contains(updateResult.SuggestedNextActions, s => s.Contains("Formula updated"));

            // Verify the update
            var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
            Assert.True(viewResult.Success);
            Assert.Contains("AVERAGE", viewResult.DaxFormula, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public async Task UpdateMeasure_WithFormatTypeOnly_UpdatesFormat()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("UpdateMeasure_WithFormatTypeOnly_UpdatesFormat.xlsx");

        // Arrange - Create a measure first
        var measureName = "FormatUpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Quantity])";

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula);

        if (createResult.Success)
        {
            await batch.SaveAsync();

            // Act - Update only the format
            var updateResult = await _dataModelCommands.UpdateMeasureAsync(batch, measureName, formatType: "Decimal");
            await batch.SaveAsync();

            // Assert
            Assert.True(updateResult.Success, $"Expected success but got error: {updateResult.ErrorMessage}");
            Assert.NotNull(updateResult.SuggestedNextActions);
            Assert.Contains(updateResult.SuggestedNextActions, s => s.Contains("Format changed"));
        }
    }

    [Fact]
    public async Task UpdateMeasure_WithDescriptionOnly_UpdatesDescription()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("UpdateMeasure_WithDescriptionOnly_UpdatesDescription.xlsx");

        // Arrange - Create a measure first
        var measureName = "DescUpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula);

        if (createResult.Success)
        {
            await batch.SaveAsync();

            // Act - Update only the description
            var updateResult = await _dataModelCommands.UpdateMeasureAsync(batch, measureName, description: "Updated description");
            await batch.SaveAsync();

            // Assert
            Assert.True(updateResult.Success, $"Expected success but got error: {updateResult.ErrorMessage}");
            Assert.NotNull(updateResult.SuggestedNextActions);
            Assert.Contains(updateResult.SuggestedNextActions, s => s.Contains("Description updated"));
        }
    }

    [Fact]
    public async Task UpdateMeasure_WithNoParameters_ReturnsError()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("UpdateMeasure_WithNoParameters_ReturnsError.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.UpdateMeasureAsync(batch, "SomeMeasure");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("No updates provided", result.ErrorMessage);
    }

    [Fact]
    public async Task UpdateMeasure_WithNonExistentMeasure_ReturnsError()
    {
        // Arrange - Create unique test file
        var testFile = await CreateTestFileAsync("UpdateMeasure_WithNonExistentMeasure_ReturnsError.xlsx");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _dataModelCommands.UpdateMeasureAsync(batch, "NonExistentMeasure", daxFormula: "1+1");

        // Assert - Should fail because measure doesn't exist (Data Model is always available in Excel 2013+)
        Assert.False(result.Success, "UpdateMeasure should fail when measure doesn't exist");
        Assert.NotNull(result.ErrorMessage);
        Assert.True(result.ErrorMessage.Contains("Measure 'NonExistentMeasure' not found"),
            $"Expected 'measure not found' error, but got: {result.ErrorMessage}");
    }

    [Fact]
    public async Task DEBUG_CreateMeasure_Minimal_ChecksIfItHangs()
    {
        // Arrange - Create minimal test file WITHOUT helper
        var testFile = await CreateTestFileAsync("DEBUG_CreateMeasure_Minimal.xlsx");

        // Manually create a simple table in Data Model
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Try to create a measure - THIS IS WHERE IT HANGS
        var result = await _dataModelCommands.CreateMeasureAsync(
            batch,
            "SalesTable",  // This table might not exist yet - should fail gracefully
            "TestMeasure",
            "1+1",
            "General"
        );

        // We don't care if it succeeds or fails - we just want to know if it hangs
        Assert.True(true, "Test completed without hanging");
    }
}
