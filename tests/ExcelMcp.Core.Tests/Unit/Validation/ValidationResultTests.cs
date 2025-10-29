using Sbroenne.ExcelMcp.Core.Models.Validation;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit.Validation;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
public class ValidationResultTests
{
    [Fact]
    public void Success_CreatesValidResult()
    {
        // Act
        var result = ValidationResult.Success();

        // Assert
        Assert.True(result.IsValid);
        Assert.Null(result.ErrorMessage);
        Assert.Null(result.ParameterName);
    }

    [Fact]
    public void Failure_CreatesInvalidResult()
    {
        // Arrange
        const string parameterName = "testParam";
        const string errorMessage = "Test error";

        // Act
        var result = ValidationResult.Failure(parameterName, errorMessage);

        // Assert
        Assert.False(result.IsValid);
        Assert.Equal(errorMessage, result.ErrorMessage);
        Assert.Equal(parameterName, result.ParameterName);
    }
}
