using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// CLI-specific tests for FileCommands - verifying argument parsing, exit codes, and CLI behavior
/// 
/// LAYER RESPONSIBILITY:
/// - ✅ Test argument validation (missing args, invalid args)
/// - ✅ Test exit code mapping (0 for success, 1 for error)
/// - ✅ Test user interaction (prompts, console output if applicable)
/// - ❌ DO NOT test Excel operations or file creation logic (that's Core's responsibility)
/// 
/// These tests verify the CLI wrapper works correctly. Business logic is tested in ExcelMcp.Core.Tests.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Files")]
[Trait("Layer", "CLI")]
public class CliFileCommandsTests : IDisposable
{
    private readonly FileCommands _cliCommands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles;

    public CliFileCommandsTests()
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
        // Note: No GC.Collect() needed here - Core's batch API handles COM cleanup properly
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
                try
                {
                    Directory.Delete(_tempDir, true);
                }
                catch
                {
                    // Best effort cleanup - test cleanup failure is non-critical
                }
            }
        }
        catch { }

        GC.SuppressFinalize(this);
    }
}
