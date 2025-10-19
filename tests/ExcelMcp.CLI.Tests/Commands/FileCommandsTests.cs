using Xunit;
using Sbroenne.ExcelMcp.Core.Commands;
using System.IO;

namespace Sbroenne.ExcelMcp.CLI.Tests.Commands;

/// <summary>
/// Integration tests for file operations including Excel workbook creation and management.
/// These tests require Excel installation and validate file manipulation commands.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Files")]
public class FileCommandsTests : IDisposable
{
    private readonly FileCommands _fileCommands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles;

    public FileCommandsTests()
    {
        _fileCommands = new FileCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_FileTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _createdFiles = new List<string>();
    }

    [Fact]
    public void CreateEmpty_WithValidPath_CreatesExcelFile()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "TestFile.xlsx");
        _createdFiles.Add(testFile);

        // Act
        var result = _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(testFile));
        
        // Verify it's a valid Excel file by checking size > 0
        var fileInfo = new FileInfo(testFile);
        Assert.True(fileInfo.Length > 0);
    }

    [Fact]
    public void CreateEmpty_WithNestedDirectory_CreatesDirectoryAndFile()
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
    public void CreateEmpty_WithInvalidArgs_ReturnsError()
    {
        // Arrange
        string invalidPath = ""; // Empty file path

        // Act
        var result = _fileCommands.CreateEmpty(invalidPath);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }

    [Fact]
    public void CreateEmpty_WithRelativePath_CreatesFileWithAbsolutePath()
    {
        // Arrange
        string relativePath = "RelativeTestFile.xlsx";
        
        // The file will be created in the current directory
        string expectedPath = Path.GetFullPath(relativePath);
        _createdFiles.Add(expectedPath);

        // Act
        var result = _fileCommands.CreateEmpty(relativePath);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(expectedPath));
    }

    [Theory]
    [InlineData("TestFile.xlsx")]
    [InlineData("TestFile.xlsm")]
    public void CreateEmpty_WithValidExtensions_CreatesFile(string fileName)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);
        _createdFiles.Add(testFile);

        // Act
        var result = _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.True(File.Exists(testFile));
    }

    [Theory]
    [InlineData("TestFile.xls")]
    [InlineData("TestFile.csv")]
    [InlineData("TestFile.txt")]
    public void CreateEmpty_WithInvalidExtensions_ReturnsError(string fileName)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);

        // Act
        var result = _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.False(File.Exists(testFile));
    }

    [Fact]
    public void CreateEmpty_WithInvalidPath_ReturnsError()
    {
        // Arrange - Use invalid characters in path
        string invalidPath = Path.Combine(_tempDir, "invalid<>file.xlsx");

        // Act
        var result = _fileCommands.CreateEmpty(invalidPath);

        // Assert
        Assert.False(result.Success);
    }

    [Fact]
    public void CreateEmpty_MultipleTimes_CreatesMultipleFiles()
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
            Assert.True(File.Exists(testFile));
        }

        // Verify all files exist
        foreach (string testFile in testFiles)
        {
            Assert.True(File.Exists(testFile));
        }
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
