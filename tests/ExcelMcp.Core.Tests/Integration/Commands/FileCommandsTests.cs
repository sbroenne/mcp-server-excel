using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Core business logic tests for FileCommands - testing Excel operations and Result objects
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test all Excel COM file operations (create, validate, etc.)
/// - ✅ Test Result object properties and error messages
/// - ✅ Test edge cases and error scenarios
/// - ✅ Test actual file system operations
/// - ❌ DO NOT test CLI argument parsing or console output (that's CLI's responsibility)
/// - ❌ DO NOT test JSON serialization (that's MCP Server's responsibility)
///
/// NOTE: ExcelSession handles all COM cleanup automatically. Tests await async operations,
/// so COM objects are fully released by the time tests complete.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Files")]
[Trait("Layer", "Core")]
public class FileCommandsTests : IDisposable
{
    private readonly FileCommands _fileCommands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles;

    public FileCommandsTests()
    {
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_FileTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _createdFiles = new List<string>();
    }

    [Fact]
    public async Task CreateEmpty_WithValidPath_ReturnsSuccessResult()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "TestFile.xlsx");
        _createdFiles.Add(testFile);

        // Act
        var result = await _fileCommands.CreateEmptyAsync(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal("create-empty", result.Action);
        Assert.NotNull(result.FilePath);
        Assert.True(File.Exists(testFile));

        // Verify it's a valid Excel file by checking size > 0
        var fileInfo = new FileInfo(testFile);
        Assert.True(fileInfo.Length > 0);
    }

    [Fact]
    public async Task CreateEmpty_WithNestedDirectory_CreatesDirectoryAndReturnsSuccess()
    {
        // Arrange
        string nestedDir = Path.Combine(_tempDir, "Nested", "Subdirectory");
        string testFile = Path.Combine(nestedDir, "NestedFile.xlsx");
        _createdFiles.Add(testFile);

        // Act
        var result = await _fileCommands.CreateEmptyAsync(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.True(Directory.Exists(nestedDir));
        Assert.True(File.Exists(testFile));
    }

    [Fact]
    public async Task CreateEmpty_WithEmptyPath_ReturnsErrorResult()
    {
        // Arrange
        string invalidPath = "";

        // Act
        var result = await _fileCommands.CreateEmptyAsync(invalidPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Equal("create-empty", result.Action);
    }

    [Fact]
    public async Task CreateEmpty_WithRelativePath_ConvertsToAbsoluteAndReturnsSuccess()
    {
        // Arrange
        string relativePath = "TestFile.xlsx";
        string expectedAbsPath = Path.GetFullPath(relativePath);
        _createdFiles.Add(expectedAbsPath);

        // Act
        var result = await _fileCommands.CreateEmptyAsync(relativePath);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(expectedAbsPath));
        Assert.Equal(expectedAbsPath, Path.GetFullPath(result.FilePath!));
    }

    [Theory]
    [InlineData("TestFile.xlsx")]
    [InlineData("TestFile.xlsm")]
    public async Task CreateEmpty_WithValidExtensions_ReturnsSuccessResult(string fileName)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);
        _createdFiles.Add(testFile);

        // Act
        var result = await _fileCommands.CreateEmptyAsync(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.True(File.Exists(testFile));
    }

    [Theory]
    [InlineData("TestFile.xls")]
    [InlineData("TestFile.csv")]
    [InlineData("TestFile.txt")]
    public async Task CreateEmpty_WithInvalidExtensions_ReturnsErrorResult(string fileName)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);

        // Act
        var result = await _fileCommands.CreateEmptyAsync(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("extension", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(testFile));
    }

    [Fact]
    public async Task CreateEmpty_WithInvalidPath_ReturnsErrorResult()
    {
        // Arrange - Use invalid characters in path
        string invalidPath = Path.Combine(_tempDir, "invalid<>file.xlsx");

        // Act
        var result = await _fileCommands.CreateEmptyAsync(invalidPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public async Task CreateEmpty_MultipleTimes_ReturnsSuccessForEachFile()
    {
        // Arrange
        string[] testFiles = {
            Path.Combine(_tempDir, "File1.xlsx"),
            Path.Combine(_tempDir, "File2.xlsx"),
            Path.Combine(_tempDir, "File3.xlsx")
        };

        _createdFiles.AddRange(testFiles);

        // Act & Assert
        foreach (string testFile in testFiles)
        {
            var result = await _fileCommands.CreateEmptyAsync(testFile);

            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.True(File.Exists(testFile));
        }
    }

    [Fact]
    public async Task CreateEmpty_FileAlreadyExists_WithoutOverwrite_ReturnsError()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "ExistingFile.xlsx");
        _createdFiles.Add(testFile);

        // Create file first
        var firstResult = await _fileCommands.CreateEmptyAsync(testFile);
        Assert.True(firstResult.Success);

        // Act - Try to create again without overwrite flag
        var result = await _fileCommands.CreateEmptyAsync(testFile, overwriteIfExists: false);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CreateEmpty_FileAlreadyExists_WithOverwrite_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "OverwriteFile.xlsx");
        _createdFiles.Add(testFile);

        // Create file first
        var firstResult = await _fileCommands.CreateEmptyAsync(testFile);
        Assert.True(firstResult.Success);

        // Get original file info
        var originalInfo = new FileInfo(testFile);
        var originalTime = originalInfo.LastWriteTime;

        // Wait a bit to ensure different timestamp
        System.Threading.Thread.Sleep(100);

        // Act - Overwrite
        var result = await _fileCommands.CreateEmptyAsync(testFile, overwriteIfExists: true);

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);

        // Verify file was overwritten (new timestamp)
        var newInfo = new FileInfo(testFile);
        Assert.True(newInfo.LastWriteTime > originalTime);
    }


    public void Dispose()
    {
        // Clean up test files
        // NOTE: ExcelSession handles all COM cleanup synchronously, so files are already released
        // by the time tests complete. Just need basic file deletion.
        try
        {
            // Brief delay for any pending async operations
            System.Threading.Thread.Sleep(200);

            // Delete individual files first
            foreach (string file in _createdFiles)
            {
                try
                {
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                    // Best effort cleanup
                }
            }

            // Then delete the temp directory
            if (Directory.Exists(_tempDir))
            {
                try
                {
                    Directory.Delete(_tempDir, true);
                }
                catch
                {
                    // Best effort cleanup - don't fail tests if cleanup fails
                }
            }
        }
        catch
        {
            // Best effort cleanup - don't fail tests if cleanup fails
        }

        GC.SuppressFinalize(this);
    }
}
