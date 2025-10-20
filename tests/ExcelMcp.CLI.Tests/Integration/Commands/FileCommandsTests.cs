using Xunit;
using Sbroenne.ExcelMcp.CLI.Commands;
using System.IO;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// Tests for CLI FileCommands - verifying CLI-specific behavior (formatting, user interaction)
/// These tests focus on the presentation layer, not the data layer
/// Core data logic is tested in ExcelMcp.Core.Tests
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Files")]
[Trait("Layer", "CLI")]
public class FileCommandsTests : IDisposable
{
    private readonly FileCommands _cliCommands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles;

    public FileCommandsTests()
    {
        _cliCommands = new FileCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_FileTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _createdFiles = new List<string>();
    }

    [Fact]
    public void CreateEmpty_WithValidPath_ReturnsZeroAndCreatesFile()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "TestFile.xlsx");
        string[] args = { "create-empty", testFile };
        _createdFiles.Add(testFile);

        // Act - CLI wraps Core and returns int exit code
        int exitCode = _cliCommands.CreateEmpty(args);

        // Assert - CLI returns 0 for success
        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(testFile));
    }

    [Fact]
    public void CreateEmpty_WithMissingArguments_ReturnsOneAndDoesNotCreateFile()
    {
        // Arrange
        string[] args = { "create-empty" }; // Missing file path

        // Act
        int exitCode = _cliCommands.CreateEmpty(args);

        // Assert - CLI returns 1 for error
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void CreateEmpty_WithInvalidExtension_ReturnsOneAndDoesNotCreateFile()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "InvalidFile.txt");
        string[] args = { "create-empty", testFile };

        // Act
        int exitCode = _cliCommands.CreateEmpty(args);

        // Assert
        Assert.Equal(1, exitCode);
        Assert.False(File.Exists(testFile));
    }

    [Theory]
    [InlineData("TestFile.xlsx")]
    [InlineData("TestFile.xlsm")]
    public void CreateEmpty_WithValidExtensions_ReturnsZero(string fileName)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);
        string[] args = { "create-empty", testFile };
        _createdFiles.Add(testFile);

        // Act
        int exitCode = _cliCommands.CreateEmpty(args);

        // Assert
        Assert.Equal(0, exitCode);
        Assert.True(File.Exists(testFile));
    }

    public void Dispose()
    {
        // Clean up test files
        try
        {
            System.Threading.Thread.Sleep(500);
            
            foreach (string file in _createdFiles)
            {
                try
                {
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                }
                catch { }
            }
            
            if (Directory.Exists(_tempDir))
            {
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, true);
                        break;
                    }
                    catch (IOException)
                    {
                        if (i == 2) throw;
                        System.Threading.Thread.Sleep(1000);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
        }
        catch { }
        
        GC.SuppressFinalize(this);
    }
}
