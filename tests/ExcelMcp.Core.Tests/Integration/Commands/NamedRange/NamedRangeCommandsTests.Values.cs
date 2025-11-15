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
    public async Task Set_ExistingParameter_UpdatesValue()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Set_ExistingParameter_UpdatesValue), _tempDir);

        // Act - Use single batch for create, set, and verify
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create parameter first
        var createResult = _parameterCommands.Create(batch, "SetTestParam", "Sheet1!A1");
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");

        // Set the parameter value
        var result = _parameterCommands.Set(batch, "SetTestParam", "TestValue");
        Assert.True(result.Success, $"Failed to set parameter: {result.ErrorMessage}");

        // Verify the parameter value was actually set by reading it back
        var getResult = _parameterCommands.Get(batch, "SetTestParam");
        Assert.True(getResult.Success, $"Failed to get parameter: {getResult.ErrorMessage}");
        Assert.Equal("TestValue", getResult.Value?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Get_ExistingParameter_ReturnsValue()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Get_ExistingParameter_ReturnsValue), _tempDir);
        string testValue = "Integration Test Value";

        // Act - Use single batch for create, set, and get
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create and set parameter value
        var createResult = _parameterCommands.Create(batch, "GetTestParam", "Sheet1!A1");
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");

        var setResult = _parameterCommands.Set(batch, "GetTestParam", testValue);
        Assert.True(setResult.Success, $"Failed to set parameter: {setResult.ErrorMessage}");

        // Get the parameter value
        var getResult = _parameterCommands.Get(batch, "GetTestParam");

        // Assert
        Assert.True(getResult.Success, $"Failed to get parameter: {getResult.ErrorMessage}");
        Assert.Equal(testValue, getResult.Value?.ToString());
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Get_WithNonExistentParameter_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFile(
            nameof(NamedRangeCommandsTests), nameof(Get_WithNonExistentParameter_ReturnsError), _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _parameterCommands.Get(batch, "NonExistentParam");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }
}
