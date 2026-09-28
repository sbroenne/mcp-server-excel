using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for Parameter value operations (get, set)
/// </summary>
public sealed partial class PersistentServiceNamedRangeTests
{
    /// <inheritdoc/>
    [Fact]
    public void Set_ExistingParameter_UpdatesValue()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Create parameter first
        _parameterCommands.Create(batch, paramName, cellRef);
        _fixture.RegisterNamedRangeForCleanup(paramName);

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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var paramName = CreateUniqueNamedRangeName();
        var cellRef = $"'{sheetName}'!A1";

        // Create and set parameter value
        _parameterCommands.Create(batch, paramName, cellRef);
        _fixture.RegisterNamedRangeForCleanup(paramName);
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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var exception = Assert.Throws<InvalidOperationException>(
            () => _parameterCommands.Read(batch, $"NonExistent_{Guid.NewGuid():N}"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}



