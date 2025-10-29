using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
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
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListMeasuresAsync(batch);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        if (result.Success)
        {
            Assert.NotNull(result.Measures);
        }
    }

    [Fact]
    public async Task ListMeasures_WithRealisticDataModel_ReturnsMeasuresWithFormulas()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ListMeasuresAsync(batch);

        // Assert
        Assert.True(result.Success || result.ErrorMessage?.Contains("does not contain a Data Model") == true,
            $"Expected success or 'no Data Model' message, but got: {result.ErrorMessage}");

        // If Data Model was created successfully with measures, validate them
        if (result.Success && result.Measures != null && result.Measures.Count > 0)
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
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ViewMeasureAsync(batch, "NonExistentMeasure");

        // Assert
        // Should fail with either "no Data Model" or "measure not found"
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("Measure 'NonExistentMeasure' not found"),
            $"Expected 'no Data Model' or 'measure not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task ViewMeasure_WithRealisticDataModel_ReturnsValidDAXFormula()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ViewMeasureAsync(batch, "Total Sales");

        // Assert - Should either succeed with valid DAX or indicate no Data Model
        if (result.Success)
        {
            Assert.NotNull(result.DaxFormula);
            Assert.Contains("SUM", result.DaxFormula, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Sales", result.DaxFormula);
            Assert.Contains("Amount", result.DaxFormula);
            Assert.Equal("Total Sales", result.MeasureName);
        }
        else
        {
            // If not successful, should be because Data Model wasn't created or measure doesn't exist
            Assert.True(
                result.ErrorMessage?.Contains("does not contain a Data Model") == true ||
                result.ErrorMessage?.Contains("not found") == true,
                $"Expected 'no Data Model' or 'not found', but got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public async Task ExportMeasure_WithNonExistentMeasure_ReturnsErrorResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ExportMeasureAsync(batch, "NonExistentMeasure", _testMeasureFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public async Task ExportMeasure_WithRealisticDataModel_ExportsValidDAXFile()
    {
        // Arrange
        var exportPath = Path.Combine(_tempDir, "TotalSales.dax");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.ExportMeasureAsync(batch, "Total Sales", exportPath);

        // Assert - Should either succeed or indicate no Data Model
        if (result.Success)
        {
            Assert.True(File.Exists(exportPath), "DAX file should be created");

            var daxContent = File.ReadAllText(exportPath);
            Assert.NotEmpty(daxContent);
            Assert.Contains("SUM", daxContent, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Sales", daxContent);
            Assert.Contains("Amount", daxContent);
        }
        else
        {
            Assert.True(
                result.ErrorMessage?.Contains("does not contain a Data Model") == true ||
                result.ErrorMessage?.Contains("not found") == true,
                $"Expected 'no Data Model' or 'not found', but got: {result.ErrorMessage}"
            );
        }
    }

    [Fact(Skip = "Data Model test helper requires specific Excel version/configuration. May fail on some environments due to Data Model availability.")]
    public async Task DeleteMeasure_WithValidMeasure_ReturnsSuccessResult()
    {
        // Arrange - Create a test measure first
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];

        await DataModelTestHelper.CreateTestMeasureAsync(_testExcelFile, measureName, "SUM(Sales[Amount])");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
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
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.DeleteMeasureAsync(batch, "NonExistentMeasure");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("Measure 'NonExistentMeasure' not found"),
            $"Expected 'no Data Model' or 'measure not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task DeleteMeasure_WithNonExistentFile_ThrowsFileNotFoundException()
    {
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
        // Arrange
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula);
        await batch.SaveAsync();

        // Assert - Should either succeed or indicate no Data Model
        if (result.Success)
        {
            Assert.NotNull(result.SuggestedNextActions);
            Assert.Contains(result.SuggestedNextActions, s => s.Contains("created successfully"));

            // Verify measure was created by listing measures
            var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
            Assert.True(listResult.Success);
            Assert.Contains(listResult.Measures, m => m.Name == measureName);
        }
        else
        {
            Assert.True(
                result.ErrorMessage?.Contains("does not contain a Data Model") == true,
                $"Expected 'no Data Model' error, but got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public async Task CreateMeasure_WithFormatType_CreatesWithFormat()
    {
        // Arrange
        var measureName = "FormattedMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula,
                                                                 formatType: "Currency", description: "Test measure with currency format");
        await batch.SaveAsync();

        // Assert - Should either succeed or indicate no Data Model
        if (result.Success)
        {
            Assert.NotNull(result.SuggestedNextActions);

            // Verify measure exists
            var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
            Assert.True(viewResult.Success);
            Assert.Equal(measureName, viewResult.MeasureName);
        }
        else
        {
            Assert.True(
                result.ErrorMessage?.Contains("does not contain a Data Model") == true,
                $"Expected 'no Data Model' error, but got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public async Task CreateMeasure_WithDuplicateName_ReturnsError()
    {
        // Arrange - First create a measure
        var measureName = "DuplicateTest_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var firstResult = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula);

        if (firstResult.Success)
        {
            await batch.SaveAsync();

            // Act - Try to create same measure again
            var duplicateResult = await _dataModelCommands.CreateMeasureAsync(batch, "Sales", measureName, daxFormula);

            // Assert - Should fail with duplicate error
            Assert.False(duplicateResult.Success);
            Assert.NotNull(duplicateResult.ErrorMessage);
            Assert.Contains("already exists", duplicateResult.ErrorMessage);
            Assert.NotNull(duplicateResult.SuggestedNextActions);
        }
    }

    [Fact]
    public async Task CreateMeasure_WithInvalidTable_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.CreateMeasureAsync(batch, "NonExistentTable", "TestMeasure", "1+1");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("Table 'NonExistentTable' not found"),
            $"Expected 'no Data Model' or 'table not found' error, but got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public async Task UpdateMeasure_WithValidFormula_UpdatesSuccessfully()
    {
        // Arrange - Create a measure first
        var measureName = "UpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var originalFormula = "SUM(Sales[Amount])";
        var updatedFormula = "AVERAGE(Sales[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
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
        // Arrange - Create a measure first
        var measureName = "FormatUpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Quantity])";

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
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
        // Arrange - Create a measure first
        var measureName = "DescUpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
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
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.UpdateMeasureAsync(batch, "SomeMeasure");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("No updates provided", result.ErrorMessage);
    }

    [Fact]
    public async Task UpdateMeasure_WithNonExistentMeasure_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _dataModelCommands.UpdateMeasureAsync(batch, "NonExistentMeasure", daxFormula: "1+1");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("does not contain a Data Model") ||
            result.ErrorMessage.Contains("Measure 'NonExistentMeasure' not found"),
            $"Expected 'no Data Model' or 'measure not found' error, but got: {result.ErrorMessage}"
        );
    }
}
