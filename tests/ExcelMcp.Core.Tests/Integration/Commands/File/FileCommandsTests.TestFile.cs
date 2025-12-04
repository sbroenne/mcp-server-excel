using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.File;

/// <summary>
/// Tests for FileCommands TestFile operation
/// </summary>
public partial class FileCommandsTests
{
    [Fact]
    public void Test_ExistingValidFile_ReturnsSuccess()
    {
        // Arrange - Create a valid file
        var testFile = _fixture.CreateTestFile();

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.True(info.IsValid);
        Assert.Equal(".xlsx", info.Extension);
        Assert.True(info.Size > 0);
        Assert.Null(info.Message);
    }
    [Fact]
    public void Test_NonExistent_ReturnsFailure()
    {
        // Arrange
        string testFile = Path.Join(_fixture.TempDir, $"NonExistent_{Guid.NewGuid():N}.xlsx");

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.False(info.Exists);
        Assert.False(info.IsValid);
        Assert.NotNull(info.Message);
        Assert.Contains("not found", info.Message, StringComparison.OrdinalIgnoreCase);
    }
    [Theory]
    [InlineData("TestFile.xls", ".xls")]
    [InlineData("TestFile.csv", ".csv")]
    [InlineData("TestFile.txt", ".txt")]
    public void Test_InvalidExtension_ReturnsFailure(string fileName, string expectedExt)
    {
        // Arrange
        string testFile = Path.Join(_fixture.TempDir, $"{Guid.NewGuid():N}_{fileName}");

        // Create file with invalid extension
        System.IO.File.WriteAllText(testFile, "test content");

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.Equal(expectedExt, info.Extension);
        Assert.NotNull(info.Message);
        Assert.Contains("Invalid file extension", info.Message, StringComparison.OrdinalIgnoreCase);
    }
}
