using System.Runtime.InteropServices;
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
    [Fact]
    public void ReadStyleIndentLevel_NativeNullConversion_ReturnsNull()
    {
        Assert.Null(ComUtilities.ReadStyleIndentLevel(
            () => throw new InvalidCastException("Native style indentation is null.", unchecked((int)0x80004002))));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    public void ReadStyleIndentLevel_ExplicitValue_PreservesValue(int value)
    {
        Assert.Equal(value, ComUtilities.ReadStyleIndentLevel(() => value));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Usage", "CA2201",
        Justification = "Injects exact native error types and codes to prove the single style getter does not hide unrelated failures.")]
    public void ReadStyleIndentLevel_UnrelatedFailure_PropagatesOriginal(int failureKind)
    {
        Exception expected = failureKind switch
        {
            0 => new COMException("Interface failure, not a null conversion.", unchecked((int)0x80004002)),
            1 => new COMException("Excel operation failed.", unchecked((int)0x800A03EC)),
            2 => new InvalidCastException("Different conversion failure.", unchecked((int)0x80070057)),
            3 => new OperationCanceledException("Cancelled."),
            _ => new InvalidOperationException("Invalid state.")
        };
        Assert.Same(expected, Record.Exception(() =>
            ComUtilities.ReadStyleIndentLevel(() => throw expected)));
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
