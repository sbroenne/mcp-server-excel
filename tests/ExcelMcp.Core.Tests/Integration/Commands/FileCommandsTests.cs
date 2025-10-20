using Xunit;
using Sbroenne.ExcelMcp.Core.Commands;
using System.IO;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Unit tests for Core FileCommands - testing data layer without UI
/// These tests verify that Core returns correct Result objects
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Files")]
[Trait("Layer", "Core")]
public class CoreFileCommandsTests : IDisposable
{
    private readonly FileCommands _fileCommands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles;

    public CoreFileCommandsTests()
    {
        _fileCommands = new FileCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_FileTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _createdFiles = new List<string>();
    }

    [Fact]
    public void CreateEmpty_WithValidPath_ReturnsSuccessResult()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "TestFile.xlsx");
        _createdFiles.Add(testFile);

        // Act
        var result = _fileCommands.CreateEmpty(testFile);

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
    public void CreateEmpty_WithNestedDirectory_CreatesDirectoryAndReturnsSuccess()
    {
        // Arrange
        string nestedDir = Path.Combine(_tempDir, "nested", "deep", "path");
        string testFile = Path.Combine(nestedDir, "TestFile.xlsx");
        _createdFiles.Add(testFile);

        // Act
        var result = _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.True(Directory.Exists(nestedDir));
        Assert.True(File.Exists(testFile));
    }

    [Fact]
    public void CreateEmpty_WithEmptyPath_ReturnsErrorResult()
    {
        // Arrange
        string invalidPath = "";

        // Act
        var result = _fileCommands.CreateEmpty(invalidPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Equal("create-empty", result.Action);
    }

    [Fact]
    public void CreateEmpty_WithRelativePath_ConvertsToAbsoluteAndReturnsSuccess()
    {
        // Arrange
        string relativePath = "RelativeTestFile.xlsx";
        string expectedPath = Path.GetFullPath(relativePath);
        _createdFiles.Add(expectedPath);

        // Act
        var result = _fileCommands.CreateEmpty(relativePath);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(expectedPath));
        Assert.Equal(expectedPath, Path.GetFullPath(result.FilePath!));
    }

    [Theory]
    [InlineData("TestFile.xlsx")]
    [InlineData("TestFile.xlsm")]
    public void CreateEmpty_WithValidExtensions_ReturnsSuccessResult(string fileName)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);
        _createdFiles.Add(testFile);

        // Act
        var result = _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.True(File.Exists(testFile));
    }

    [Theory]
    [InlineData("TestFile.xls")]
    [InlineData("TestFile.csv")]
    [InlineData("TestFile.txt")]
    public void CreateEmpty_WithInvalidExtensions_ReturnsErrorResult(string fileName)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);

        // Act
        var result = _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("extension", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(testFile));
    }

    [Fact]
    public void CreateEmpty_WithInvalidPath_ReturnsErrorResult()
    {
        // Arrange - Use invalid characters in path
        string invalidPath = Path.Combine(_tempDir, "invalid<>file.xlsx");

        // Act
        var result = _fileCommands.CreateEmpty(invalidPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void CreateEmpty_MultipleTimes_ReturnsSuccessForEachFile()
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
            var result = _fileCommands.CreateEmpty(testFile);
            
            Assert.True(result.Success);
            Assert.Null(result.ErrorMessage);
            Assert.True(File.Exists(testFile));
        }
    }

    [Fact]
    public void CreateEmpty_FileAlreadyExists_WithoutOverwrite_ReturnsError()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "ExistingFile.xlsx");
        _createdFiles.Add(testFile);
        
        // Create file first
        var firstResult = _fileCommands.CreateEmpty(testFile);
        Assert.True(firstResult.Success);

        // Act - Try to create again without overwrite flag
        var result = _fileCommands.CreateEmpty(testFile, overwriteIfExists: false);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CreateEmpty_FileAlreadyExists_WithOverwrite_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "OverwriteFile.xlsx");
        _createdFiles.Add(testFile);
        
        // Create file first
        var firstResult = _fileCommands.CreateEmpty(testFile);
        Assert.True(firstResult.Success);
        
        // Get original file info
        var originalInfo = new FileInfo(testFile);
        var originalTime = originalInfo.LastWriteTime;
        
        // Wait a bit to ensure different timestamp
        System.Threading.Thread.Sleep(100);

        // Act - Overwrite
        var result = _fileCommands.CreateEmpty(testFile, overwriteIfExists: true);

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
        try
        {
            // Wait a bit for Excel to fully release files
            System.Threading.Thread.Sleep(500);
            
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
                // Try to delete directory multiple times if needed
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, true);
                        break;
                    }
                    catch (IOException)
                    {
                        if (i == 2) throw; // Last attempt failed
                        System.Threading.Thread.Sleep(1000);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
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
