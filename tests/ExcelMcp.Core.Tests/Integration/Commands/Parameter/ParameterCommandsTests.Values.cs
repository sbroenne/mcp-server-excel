using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Parameter;

/// <summary>
/// Tests for Parameter value operations (get, set)
/// </summary>
public partial class ParameterCommandsTests
{
    [Fact]
    public async Task Set_WithValidParameter_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ParameterCommandsTests), nameof(Set_WithValidParameter_ReturnsSuccess), _tempDir);

        // Act - Use single batch for create, set, and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create parameter first
        var createResult = await _parameterCommands.CreateAsync(batch, "SetTestParam", "Sheet1!A1");
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");
        
        // Set the parameter value
        var result = await _parameterCommands.SetAsync(batch, "SetTestParam", "TestValue");
        Assert.True(result.Success, $"Failed to set parameter: {result.ErrorMessage}");

        // Verify the parameter value was actually set by reading it back
        var getResult = await _parameterCommands.GetAsync(batch, "SetTestParam");
        Assert.True(getResult.Success, $"Failed to get parameter: {getResult.ErrorMessage}");
        Assert.Equal("TestValue", getResult.Value?.ToString());

        await batch.SaveAsync();
    }

    [Fact]
    public async Task Get_WithValidParameter_ReturnsValue()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ParameterCommandsTests), nameof(Get_WithValidParameter_ReturnsValue), _tempDir);
        string testValue = "Integration Test Value";

        // Act - Use single batch for create, set, and get
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create and set parameter value
        var createResult = await _parameterCommands.CreateAsync(batch, "GetTestParam", "Sheet1!A1");
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");
        
        var setResult = await _parameterCommands.SetAsync(batch, "GetTestParam", testValue);
        Assert.True(setResult.Success, $"Failed to set parameter: {setResult.ErrorMessage}");

        // Get the parameter value
        var getResult = await _parameterCommands.GetAsync(batch, "GetTestParam");

        // Assert
        Assert.True(getResult.Success, $"Failed to get parameter: {getResult.ErrorMessage}");
        Assert.Equal(testValue, getResult.Value?.ToString());

        await batch.SaveAsync();
    }

    [Fact]
    public async Task Get_WithNonExistentParameter_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(ParameterCommandsTests), nameof(Get_WithNonExistentParameter_ReturnsError), _tempDir);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _parameterCommands.GetAsync(batch, "NonExistentParam");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }
}
