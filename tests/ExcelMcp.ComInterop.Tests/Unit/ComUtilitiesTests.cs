using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

/// <summary>
/// Unit tests for ComUtilities helper methods.
/// These tests verify the low-level COM interop utility functions.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
[Trait("RequiresExcel", "false")]
public class ComUtilitiesTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RequiredMetadata_WithNullObject_ThrowsArgumentNullException(bool readCount)
    {
        var exception = Assert.Throws<ArgumentNullException>(() =>
        {
            if (readCount)
                ComUtilities.SafeGetInt(null, "Count");
            else
                ComUtilities.SafeGetString(null, "Name");
        });
        Assert.Equal("obj", exception.ParamName);
    }

    [Fact]
    public void Release_WithNullObject_DoesNotThrow()
    {
        // Arrange
        object? nullObject = null;

        // Act & Assert - Should not throw
        ComUtilities.Release(ref nullObject);
        Assert.Null(nullObject);
    }

    [Fact]
    public void Release_WithNonComObject_SetsToNull()
    {
        // Arrange
        string? testObject = "test";

        // Act
        ComUtilities.Release(ref testObject);

        // Assert
        Assert.Null(testObject);
    }

    [Fact]
    public void Release_MultipleCallsOnSameReference_DoesNotThrow()
    {
        // Arrange
        object? obj = new object();

        // Act & Assert - Multiple releases should be safe
        ComUtilities.Release(ref obj);
        Assert.Null(obj);

        ComUtilities.Release(ref obj); // Second call on null
        Assert.Null(obj);
    }
}

