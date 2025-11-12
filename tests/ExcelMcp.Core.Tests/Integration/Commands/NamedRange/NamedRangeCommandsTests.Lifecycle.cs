using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange;

/// <summary>
/// Tests for Parameter lifecycle operations (list, create, delete)
/// </summary>
public partial class NamedRangeCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public async Task List_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(List_EmptyWorkbook_ReturnsEmptyList), _tempDir);

        // Act
        var result = await _parameterCommands.ListAsync(testFile);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.NamedRanges);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Create_ValidNameAndReference_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Create_ValidNameAndReference_ReturnsSuccess), _tempDir);

        // Act
        var result = await _parameterCommands.CreateAsync(testFile, "TestParam", "Sheet1!A1");

        // Assert
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");

        // Verify the parameter was actually created by listing parameters
        var listResult = await _parameterCommands.ListAsync(testFile);
        Assert.True(listResult.Success, $"Failed to list parameters: {listResult.ErrorMessage}");
        Assert.Contains(listResult.NamedRanges, p => p.Name == "TestParam");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Delete_ExistingParameter_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Delete_ExistingParameter_ReturnsSuccess), _tempDir);

        // Create parameter first
        var createResult = await _parameterCommands.CreateAsync(testFile, "DeleteTestParam", "Sheet1!A1");
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");

        // Act - Delete the parameter
        var result = await _parameterCommands.DeleteAsync(testFile, "DeleteTestParam");
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Assert - Verify the parameter was actually deleted by checking it's not in the list
        var listResult = await _parameterCommands.ListAsync(testFile);
        Assert.True(listResult.Success, $"Failed to list parameters: {listResult.ErrorMessage}");
        Assert.DoesNotContain(listResult.NamedRanges, p => p.Name == "DeleteTestParam");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task List_WithNonExistentFile_ReturnsError()
    {
        // Act
        var result = await _parameterCommands.ListAsync("nonexistent.xlsx");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }
}
