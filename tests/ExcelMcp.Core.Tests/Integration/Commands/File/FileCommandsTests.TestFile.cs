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

    [Fact]
    public void Test_LegacyXls_ReturnsOpenCompatibilityGuidance()
    {
        string testFile = Path.Join(_fixture.TempDir, $"{Guid.NewGuid():N}_TestFile.xls");
        System.IO.File.WriteAllText(testFile, "test content");

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.Equal(".xls", info.Extension);
        Assert.NotNull(info.Message);
        Assert.Contains("can be opened", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("strict file-test policy", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_FileExceedingOneGiB_ReturnsFailure()
    {
        string testFile = Path.Join(_fixture.TempDir, $"{Guid.NewGuid():N}.xlsx");
        using (var stream = new FileStream(testFile, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            stream.SetLength((1024L * 1024 * 1024) + 1);
        }

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.Contains("maximum size", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_PathExceeding32767Characters_ReturnsFailure()
    {
        string testFile = Path.Join(_fixture.TempDir, new string('x', 32768));

        var info = _fileCommands.Test(testFile);

        Assert.False(info.Exists);
        Assert.False(info.IsValid);
        Assert.Contains("maximum path length", info.Message, StringComparison.OrdinalIgnoreCase);
    }
}




