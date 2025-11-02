using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.NamedRange;

/// <summary>
/// Tests for Parameter lifecycle operations (list, create, delete)
/// </summary>
public partial class NamedRangeCommandsTests
{
    [Fact]
    public async Task List_WithValidFile_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(List_WithValidFile_ReturnsSuccess), _tempDir);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _parameterCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.NamedRanges);
    }

    [Fact]
    public async Task Create_WithValidParameter_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Create_WithValidParameter_ReturnsSuccess), _tempDir);

        // Act - Use single batch for create and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _parameterCommands.CreateAsync(batch, "TestParam", "Sheet1!A1");

        // Assert
        Assert.True(result.Success, $"Create failed: {result.ErrorMessage}");

        // Verify the parameter was actually created by listing parameters
        var listResult = await _parameterCommands.ListAsync(batch);
        Assert.True(listResult.Success, $"Failed to list parameters: {listResult.ErrorMessage}");
        Assert.Contains(listResult.NamedRanges, p => p.Name == "TestParam");
    }

    [Fact]
    public async Task Delete_WithValidParameter_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(NamedRangeCommandsTests), nameof(Delete_WithValidParameter_ReturnsSuccess), _tempDir);

        // Act - Use single batch for create, delete, and verify
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create parameter first
        var createResult = await _parameterCommands.CreateAsync(batch, "DeleteTestParam", "Sheet1!A1");
        Assert.True(createResult.Success, $"Failed to create parameter: {createResult.ErrorMessage}");
        
        // Delete the parameter
        var result = await _parameterCommands.DeleteAsync(batch, "DeleteTestParam");
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Verify the parameter was actually deleted by checking it's not in the list
        var listResult = await _parameterCommands.ListAsync(batch);
        Assert.True(listResult.Success, $"Failed to list parameters: {listResult.ErrorMessage}");
        Assert.DoesNotContain(listResult.NamedRanges, p => p.Name == "DeleteTestParam");
    }

    [Fact]
    public async Task List_WithNonExistentFile_ReturnsError()
    {
        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync("nonexistent.xlsx");
        });
    }
}
