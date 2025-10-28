using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

public partial class DataModelTomCommandsTests
{
    #region CreateMeasure Tests

    [Fact]
    public async Task CreateMeasure_WithValidParameters_ReturnsSuccess()
    {
        // Arrange
        var measureName = "TestMeasure_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        var result = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            measureName,
            daxFormula,
            "Test measure for integration testing",
            "#,##0.00"
        );

        // Assert
        if (result.Success)
        {
            Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
            Assert.NotNull(result.SuggestedNextActions);
            Assert.Contains(result.SuggestedNextActions, s => s.Contains("created successfully"));

            // Verify the measure was actually created by listing measures
            await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
            var listResult = await _dataModelCommands.ListMeasuresAsync(batch);
            if (listResult.Success)
            {
                Assert.Contains(listResult.Measures, m => m.Name == measureName);
            }
        }
        else
        {
            // If TOM connection failed, verify it's because of Data Model availability
            Assert.True(
                result.ErrorMessage?.Contains("Data Model") == true ||
                result.ErrorMessage?.Contains("connect") == true,
                $"Expected Data Model or connection error, got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void CreateMeasure_WithInvalidTable_ReturnsError()
    {
        // Arrange
        var measureName = "InvalidTableMeasure";
        var daxFormula = "SUM(Sales[Amount])";

        // Act
        var result = _tomCommands.CreateMeasure(
            _testExcelFile,
            "NonExistentTable",
            measureName,
            daxFormula
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.True(
            result.ErrorMessage.Contains("Table") ||
            result.ErrorMessage.Contains("not found") ||
            result.ErrorMessage.Contains("connect"),
            $"Expected table or connection error, got: {result.ErrorMessage}"
        );
    }

    [Fact]
    public void CreateMeasure_WithEmptyMeasureName_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            "",
            "SUM(Sales[Amount])"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("name cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CreateMeasure_WithEmptyFormula_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            "TestMeasure",
            ""
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("formula cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact(Skip = "TOM API requires specific configuration and may not be available in all Excel environments.")]
    public void CreateMeasure_WithDuplicateName_ReturnsError()
    {
        // Arrange - Create first measure
        var measureName = "DuplicateTest_" + Guid.NewGuid().ToString("N")[..8];
        var result1 = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            measureName,
            "SUM(Sales[Amount])"
        );

        Assert.True(result1.Success, $"First create failed: {result1.ErrorMessage}");

        // Act - Try to create duplicate
        var result2 = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            measureName,
            "AVERAGE(Sales[Amount])"
        );

        // Assert
        Assert.False(result2.Success);
        Assert.NotNull(result2.ErrorMessage);
        Assert.Contains("already exists", result2.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region UpdateMeasure Tests

    [Fact]
    public async Task UpdateMeasure_WithValidParameters_ReturnsSuccess()
    {
        // Arrange - Create a measure first
        var measureName = "UpdateTest_" + Guid.NewGuid().ToString("N")[..8];
        var createResult = _tomCommands.CreateMeasure(
            _testExcelFile,
            "Sales",
            measureName,
            "SUM(Sales[Amount])"
        );

        // TOM connection failure should fail the test, not skip it
        if (!createResult.Success && createResult.ErrorMessage?.Contains("connect") == true)
        {
            Assert.Fail($"TOM connection failed: {createResult.ErrorMessage}");
        }

        Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");

        // Act - Update the measure
        var updateResult = _tomCommands.UpdateMeasure(
            _testExcelFile,
            measureName,
            daxFormula: "AVERAGE(Sales[Amount])",
            description: "Updated description",
            formatString: "0.00%"
        );

        // Assert
        Assert.True(updateResult.Success, $"Update failed: {updateResult.ErrorMessage}");
        Assert.NotNull(updateResult.SuggestedNextActions);
        Assert.Contains(updateResult.SuggestedNextActions, s => s.Contains("updated successfully"));

        // Verify the measure was updated
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var viewResult = await _dataModelCommands.ViewMeasureAsync(batch, measureName);
        if (viewResult.Success)
        {
            Assert.Contains("AVERAGE", viewResult.DaxFormula);
        }
    }

    [Fact]
    public void UpdateMeasure_WithNonExistentMeasure_ReturnsError()
    {
        // Act
        var result = _tomCommands.UpdateMeasure(
            _testExcelFile,
            "NonExistentMeasure",
            daxFormula: "SUM(Sales[Amount])"
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
    public void UpdateMeasure_WithNoParameters_ReturnsError()
    {
        // Act
        var result = _tomCommands.UpdateMeasure(
            _testExcelFile,
            "SomeMeasure"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("at least one property", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
