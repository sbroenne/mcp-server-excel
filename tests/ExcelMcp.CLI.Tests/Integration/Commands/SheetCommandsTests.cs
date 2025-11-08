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
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Worksheets")]
[Trait("Layer", "CLI")]
public class SheetCommandsTests
{
    private readonly SheetCommands _cliCommands;
    /// <inheritdoc/>

    public SheetCommandsTests()
    {
        _cliCommands = new SheetCommands();
    }
    /// <inheritdoc/>

    [Fact]
    public void List_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["sheet-list"]; // Missing file path

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_WithMissingArgs_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["sheet-create", "file.xlsx"]; // Missing sheet name

        // Act
        int exitCode = _cliCommands.Create(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Rename_WithMissingNewNameArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["sheet-rename", "file.xlsx", "OldName"]; // Missing new name

        // Act
        int exitCode = _cliCommands.Rename(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Copy_WithInvalidFileExtension_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["sheet-copy", "invalid.txt", "Source", "Target"];

        // Act
        int exitCode = _cliCommands.Copy(args);

        // Assert - CLI returns 1 for error (invalid file extension)
        Assert.Equal(1, exitCode);
    }
}
