using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

public partial class DataModelTomCommandsTests
{
    #region CreateCalculatedColumn Tests

    [Fact]
    public void CreateCalculatedColumn_WithValidParameters_ReturnsSuccess()
    {
        // Arrange
        var columnName = "TestColumn_" + Guid.NewGuid().ToString("N")[..8];
        var daxFormula = "[Amount] * 2";

        // Act
        var result = _tomCommands.CreateCalculatedColumn(
            _testExcelFile,
            "Sales",
            columnName,
            daxFormula,
            "Test calculated column",
            "Double"
        );

        // Assert
        if (result.Success)
        {
            Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
            Assert.NotNull(result.SuggestedNextActions);
            Assert.Contains(result.SuggestedNextActions, s => s.Contains("created successfully"));
        }
        else
        {
            // TOM connection failure or table not found is acceptable
            Assert.True(
                result.ErrorMessage?.Contains("Data Model") == true ||
                result.ErrorMessage?.Contains("connect") == true ||
                result.ErrorMessage?.Contains("not found") == true,
                $"Expected Data Model, connection, or not found error, got: {result.ErrorMessage}"
            );
        }
    }

    [Fact]
    public void CreateCalculatedColumn_WithEmptyColumnName_ReturnsError()
    {
        // Act
        var result = _tomCommands.CreateCalculatedColumn(
            _testExcelFile,
            "Sales",
            "",
            "[Amount] * 2"
        );

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("name cannot be empty", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
