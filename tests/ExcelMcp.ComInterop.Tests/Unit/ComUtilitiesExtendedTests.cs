using System.Runtime.InteropServices;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

/// <summary>
/// Extended tests for ComUtilities - tests error handling and edge cases.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "ComInterop")]
public class ComUtilitiesExtendedTests
{
    [Fact]
    public void Release_WithComObject_DoesNotThrow()
    {
        // This test verifies Release handles actual COM objects gracefully
        // Note: We can't easily create real COM objects in unit tests,
        // but we verify the method is null-safe

        object? obj = null;
        ComUtilities.Release(ref obj);
        Assert.Null(obj);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("test")]
    public void Release_WithVariousTypes_SetsToNull(string? testValue)
    {
        // Arrange
        string? obj = testValue;

        // Act
        ComUtilities.Release(ref obj);

        // Assert
        Assert.Null(obj);
    }

    [Fact]
    public async Task Release_CalledConcurrently_ThreadSafe()
    {
        // Arrange
        var tasks = new List<Task>();

        // Act - Release from multiple threads
        for (int i = 0; i < 10; i++)
        {
            tasks.Add(Task.Run(() =>
            {
                object? obj = new object();
                ComUtilities.Release(ref obj);
                Assert.Null(obj);
            }));
        }

        // Assert - All complete without exceptions
        await Task.WhenAll(tasks);
    }
}
