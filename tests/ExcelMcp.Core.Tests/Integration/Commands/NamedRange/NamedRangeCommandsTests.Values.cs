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
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Set_ExistingParameter_UpdatesValue), _tempDir);

        // Act - Use single batch for create, set, and verify
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create parameter first
        _parameterCommands.Create(batch, "SetTestParam", "Sheet1!A1");

        // Set the parameter value
        _parameterCommands.Write(batch, "SetTestParam", "TestValue");

        // Assert - Verify the parameter value was actually set by reading it back
        var namedRangeValue = _parameterCommands.Read(batch, "SetTestParam");
        Assert.Equal("TestValue", namedRangeValue.Value?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_ExistingParameter_ReturnsValue()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Get_ExistingParameter_ReturnsValue), _tempDir);
        string testValue = "Integration Test Value";

        // Act - Use single batch for create, set, and get
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create and set parameter value
        _parameterCommands.Create(batch, "GetTestParam", "Sheet1!A1");
        _parameterCommands.Write(batch, "GetTestParam", testValue);

        // Get the parameter value
        var namedRangeValue = _parameterCommands.Read(batch, "GetTestParam");

        // Assert
        Assert.Equal(testValue, namedRangeValue.Value?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_WithNonExistentParameter_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Get_WithNonExistentParameter_ThrowsException), _tempDir);

        // Act & Assert
        using var batch = ExcelSession.BeginBatch(testFile);
        var exception = Assert.Throws<InvalidOperationException>(() => _parameterCommands.Read(batch, "NonExistentParam"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
