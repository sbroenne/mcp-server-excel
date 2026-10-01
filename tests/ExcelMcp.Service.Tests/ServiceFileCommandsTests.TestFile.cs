using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for FileCommands TestFile operation
/// </summary>
public sealed partial class ServiceFileCommandsTests
{
    [Fact]
    public void Test_ExistingWorkbookRemainsStructurallyUnvalidated()
    {
        // Arrange - Create a valid file
        var testFile = _fixture.CreateTestFile();
        var lastWriteTime = File.GetLastWriteTimeUtc(testFile);

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.True(info.PreflightPassed);
        Assert.True(info.Success);
        Assert.False(info.CanOpen);
        Assert.False(info.WillOpenReadOnly);
        Assert.False(info.RequiresVisibleSession);
        Assert.Equal(".xlsx", info.Extension);
        Assert.True(info.Size > 0);
        Assert.NotNull(info.Message);
        Assert.Contains("opaque", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("file(action: 'open')", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("session open", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(lastWriteTime, File.GetLastWriteTimeUtc(testFile));
        Assert.Equal(0, _fixture.SessionCount);
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
        Assert.False(info.PreflightPassed);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.Null(info.IsError);
        Assert.NotNull(info.Message);
        Assert.Contains("not found", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_CorruptSupportedExtension_IsNotValidOrOpenable()
    {
        var testFile = Path.Join(_fixture.TempDir, $"Corrupt_{Guid.NewGuid():N}.xlsx");
        System.IO.File.WriteAllText(testFile, "not an Excel workbook");

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.True(info.PreflightPassed);
        Assert.True(info.Success);
        Assert.False(info.CanOpen);
        Assert.NotNull(info.Message);
        Assert.Contains("not inspected", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_DoesNotInferValidityFromWorkbookContents()
    {
        var testFile = Path.Join(_fixture.TempDir, $"Opaque_{Guid.NewGuid():N}.xlsx");
        System.IO.File.WriteAllBytes(testFile, [0x50, 0x4B, 0x03, 0x04]);

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.True(info.PreflightPassed);
        Assert.True(info.Success);
        Assert.False(info.CanOpen);
        Assert.Contains("not inspected", info.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Test_LockedSupportedFile_ReportsNotOpenable()
    {
        var testFile = _fixture.CreateTestFile();
        using var lockStream = new FileStream(
            testFile,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.None);

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.PreflightPassed);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.NotNull(info.Message);
        Assert.Contains("already open", info.Message, StringComparison.OrdinalIgnoreCase);
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
        Assert.False(info.PreflightPassed);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.Equal(expectedExt, info.Extension);
        Assert.NotNull(info.Message);
        Assert.Contains("Invalid file extension", info.Message, StringComparison.OrdinalIgnoreCase);
    }
}
