using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange;

/// <summary>
/// Tests for Parameter value operations (get, set)
/// </summary>
public partial class NamedRangeCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public void Set_ExistingParameter_UpdatesValue()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var paramName = NamedRangeTestsFixture.GetUniqueNamedRangeName();
        var cellRef = _fixture.GetUniqueCellReference();

        // Create parameter first
        _parameterCommands.Create(batch, paramName, cellRef);

        // Set the parameter value
        _parameterCommands.Write(batch, paramName, "TestValue");

        // Assert - Verify the parameter value was actually set by reading it back
        var namedRangeValue = _parameterCommands.Read(batch, paramName);
        Assert.Equal("TestValue", namedRangeValue.Value?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_ExistingParameter_ReturnsValue()
    {
        // Arrange
        string testValue = "Integration Test Value";
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var paramName = NamedRangeTestsFixture.GetUniqueNamedRangeName();
        var cellRef = _fixture.GetUniqueCellReference();

        // Create and set parameter value
        _parameterCommands.Create(batch, paramName, cellRef);
        _parameterCommands.Write(batch, paramName, testValue);

        // Get the parameter value
        var namedRangeValue = _parameterCommands.Read(batch, paramName);

        // Assert
        Assert.Equal(testValue, namedRangeValue.Value?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_WithNonExistentParameter_ThrowsException()
    {
        // Arrange & Act & Assert
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var exception = Assert.Throws<InvalidOperationException>(
            () => _parameterCommands.Read(batch, $"NonExistent_{Guid.NewGuid():N}"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}




