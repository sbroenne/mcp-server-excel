using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for FileCommands TestFile operation
/// </summary>
public sealed partial class ServiceFileCommandsTests
{
    [Fact]
    public void Test_ExistingValidFile_ReturnsSuccess()
    {
        // Arrange - Create a valid file
        var testFile = _fixture.CreateTestFile();
        var lastWriteTime = File.GetLastWriteTimeUtc(testFile);
        var bytes = File.ReadAllBytes(testFile);

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.True(info.IsValid);
        Assert.True(info.Success);
        Assert.True(info.CanOpen);
        Assert.False(info.WillOpenReadOnly);
        Assert.False(info.RequiresVisibleSession);
        Assert.Equal(".xlsx", info.Extension);
        Assert.True(info.Size > 0);
        Assert.Null(info.Message);
        Assert.Equal(lastWriteTime, File.GetLastWriteTimeUtc(testFile));
        Assert.Equal(bytes, File.ReadAllBytes(testFile));
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
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.Null(info.IsError);
        Assert.NotNull(info.Message);
        Assert.Contains("not found", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(testFile));
    }

    [Fact]
    public void Test_CorruptSupportedExtension_IsNotValidOrOpenable()
    {
        var testFile = Path.Join(_fixture.TempDir, $"Corrupt_{Guid.NewGuid():N}.xlsx");
        System.IO.File.WriteAllText(testFile, "not an Excel workbook");
        var bytes = File.ReadAllBytes(testFile);

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.NotNull(info.Message);
        Assert.Contains("valid Excel workbook", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("already open", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(bytes, File.ReadAllBytes(testFile));
    }

    [Theory]
    [InlineData("timeout", nameof(TimeoutException))]
    [InlineData("cancellation", nameof(OperationCanceledException))]
    public void Test_ValidationControlFlowFailure_ReturnsServiceError(
        string failure,
        string expectedExceptionType)
    {
        var testFile = _fixture.CreateTestFile();
        var bytes = File.ReadAllBytes(testFile);
        var originalHook = ExcelBatch.BeforeWorkbookOpenHook;
        ExcelBatch.BeforeWorkbookOpenHook = (_, _) => throw failure switch
        {
            "timeout" => new TimeoutException("Validation timed out."),
            "cancellation" => new OperationCanceledException("Validation cancelled."),
            _ => throw new InvalidOperationException($"Unknown failure type: {failure}")
        };

        PersistentServiceCleanupFailures.Run(() =>
        {
            var response = _fileCommands.TestRaw(testFile);

            Assert.False(response.Success);
            Assert.NotNull(response.ErrorMessage);
            Assert.True(string.IsNullOrEmpty(response.Result));
            Assert.Equal(expectedExceptionType, response.ExceptionType);
            Assert.Equal(bytes, File.ReadAllBytes(testFile));
        }, () => ExcelBatch.BeforeWorkbookOpenHook = originalHook);
        var recovered = _fileCommands.Test(testFile);
        Assert.True(recovered.Success);
        Assert.True(recovered.IsValid);
        Assert.True(recovered.CanOpen);
        Assert.Equal(bytes, File.ReadAllBytes(testFile));
    }

    [Fact]
    public void Test_LockedSupportedFile_ReportsNotOpenable()
    {
        var testFile = _fixture.CreateTestFile();
        var bytes = File.ReadAllBytes(testFile);
        using (var lockStream = new FileStream(
            testFile,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.None))
        {
            var info = _fileCommands.Test(testFile);

            Assert.True(info.Exists);
            Assert.False(info.IsValid);
            Assert.False(info.Success);
            Assert.False(info.CanOpen);
            Assert.NotNull(info.Message);
            Assert.Contains("already open", info.Message, StringComparison.OrdinalIgnoreCase);
        }
        Assert.Equal(bytes, File.ReadAllBytes(testFile));
        Assert.True(_fileCommands.Test(testFile).CanOpen);
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
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.Equal(expectedExt, info.Extension);
        Assert.NotNull(info.Message);
        Assert.Contains("Invalid file extension", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("test content", File.ReadAllText(testFile));
    }
}
