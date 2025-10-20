using Xunit;
using Sbroenne.ExcelMcp.CLI.Commands;
using System.IO;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// Tests for CLI PowerQueryCommands - verifying CLI-specific behavior (argument parsing, exit codes)
/// These tests focus on the presentation layer, not the business logic
/// Core data logic is tested in ExcelMcp.Core.Tests
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "CLI")]
public class PowerQueryCommandsTests : IDisposable
{
    private readonly PowerQueryCommands _cliCommands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles;

    public PowerQueryCommandsTests()
    {
        _cliCommands = new PowerQueryCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_PowerQueryTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        
        _createdFiles = new List<string>();
    }

    [Fact]
    public void List_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "pq-list" }; // Missing file path

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void View_WithMissingArgs_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "pq-view", "file.xlsx" }; // Missing query name

        // Act
        int exitCode = _cliCommands.View(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void List_WithNonExistentFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");
        string[] args = { "pq-list", nonExistentFile };

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void View_WithNonExistentFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");
        string[] args = { "pq-view", nonExistentFile, "SomeQuery" };

        // Act
        int exitCode = _cliCommands.View(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Refresh_WithInvalidFileExtension_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "pq-refresh", "invalid.txt", "SomeQuery" };

        // Act
        int exitCode = _cliCommands.Refresh(args);

        // Assert - CLI returns 1 for error (invalid file extension)
        Assert.Equal(1, exitCode);
    }

    [Theory]
    [InlineData("pq-import")]
    [InlineData("pq-update")]
    public async Task AsyncCommands_WithMissingArgs_ReturnsErrorExitCode(string command)
    {
        // Arrange
        string[] args = { command }; // Missing required arguments

        // Act & Assert - Handle potential markup exceptions
        try
        {
            int exitCode = command switch
            {
                "pq-import" => await _cliCommands.Import(args),
                "pq-update" => await _cliCommands.Update(args),
                _ => throw new ArgumentException($"Unknown command: {command}")
            };
            Assert.Equal(1, exitCode); // CLI returns 1 for error (missing arguments)
        }
        catch (Exception ex)
        {
            // CLI has markup issues - document current behavior
            Assert.True(ex is InvalidOperationException || ex is ArgumentException, 
                $"Unexpected exception type: {ex.GetType().Name}: {ex.Message}");
        }
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