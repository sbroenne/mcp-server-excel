using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// CLI-specific tests for SheetCommands - verifying argument parsing, exit codes, and CLI behavior
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test argument validation (missing args, invalid args)
/// - ✅ Test exit code mapping (0 for success, 1 for error)
/// - ✅ Test user interaction (prompts, console output if applicable)
/// - ❌ DO NOT test worksheet operations or Excel COM interop (that's Core's responsibility)
///
/// These tests verify the CLI wrapper works correctly. Business logic is tested in ExcelMcp.Core.Tests.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Worksheets")]
[Trait("Layer", "CLI")]
public class SheetCommandsTests : IDisposable
{
    private readonly SheetCommands _cliCommands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles;

    public SheetCommandsTests()
    {
        _cliCommands = new SheetCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_SheetTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _createdFiles = [];
    }

    [Fact]
    public void List_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "sheet-list" }; // Missing file path

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Create_WithMissingArgs_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "sheet-create", "file.xlsx" }; // Missing sheet name

        // Act
        int exitCode = _cliCommands.Create(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Delete_WithNonExistentFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");
        string[] args = { "sheet-delete", nonExistentFile, "Sheet1" };

        // Act
        int exitCode = _cliCommands.Delete(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Rename_WithMissingNewNameArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "sheet-rename", "file.xlsx", "OldName" }; // Missing new name

        // Act
        int exitCode = _cliCommands.Rename(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Copy_WithInvalidFileExtension_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "sheet-copy", "invalid.txt", "Source", "Target" };

        // Act
        int exitCode = _cliCommands.Copy(args);

        // Assert - CLI returns 1 for error (invalid file extension)
        Assert.Equal(1, exitCode);
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
