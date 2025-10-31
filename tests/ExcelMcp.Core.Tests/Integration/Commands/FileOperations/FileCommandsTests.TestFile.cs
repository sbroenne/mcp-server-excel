using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.FileOperations;

/// <summary>
/// Tests for FileCommands TestFile operation
/// </summary>
public partial class FileCommandsTests
{
    [Fact]
    public async Task TestFile_ExistingValidFile_ReturnsSuccess()
    {
        // Arrange - Create a valid file
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(FileCommandsTests), nameof(TestFile_ExistingValidFile_ReturnsSuccess), _tempDir);

        // Act
        var result = await _fileCommands.TestFileAsync(testFile);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Null(result.ErrorMessage);
        Assert.True(result.Exists);
        Assert.True(result.IsValid);
        Assert.Equal(".xlsx", result.Extension);
        Assert.True(result.Size > 0);
    }

    [Fact]
    public async Task TestFile_NonExistent_ReturnsFailure()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, $"NonExistent_{Guid.NewGuid():N}.xlsx");

        // Act
        var result = await _fileCommands.TestFileAsync(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.False(result.Exists);
        Assert.False(result.IsValid);
    }

    [Theory]
    [InlineData("TestFile.xls", ".xls")]
    [InlineData("TestFile.csv", ".csv")]
    [InlineData("TestFile.txt", ".txt")]
    public async Task TestFile_InvalidExtension_ReturnsFailure(string fileName, string expectedExt)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, $"{Guid.NewGuid():N}_{fileName}");

        // Create file with invalid extension
        File.WriteAllText(testFile, "test content");

        // Act
        var result = await _fileCommands.TestFileAsync(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Invalid file extension", result.ErrorMessage);
        Assert.True(result.Exists);
        Assert.False(result.IsValid);
        Assert.Equal(expectedExt, result.Extension);
    }
}
