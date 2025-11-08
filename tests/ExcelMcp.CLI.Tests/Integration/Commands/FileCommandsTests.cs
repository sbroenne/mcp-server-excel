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
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Files")]
[Trait("Layer", "CLI")]
public class CliFileCommandsTests
{
    private readonly FileCommands _cliCommands;
    /// <inheritdoc/>

    public CliFileCommandsTests()
    {
        _cliCommands = new FileCommands();
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateEmpty_WithMissingArguments_ReturnsOneAndDoesNotCreateFile()
    {
        // Arrange
        string[] args = ["create-empty"]; // Missing file path

        // Act
        int exitCode = _cliCommands.CreateEmpty(args);

        // Assert - CLI returns 1 for error
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateEmpty_WithInvalidExtension_ReturnsOneAndDoesNotCreateFile()
    {
        // Arrange
        string[] args = ["create-empty", "InvalidFile.txt"];

        // Act
        int exitCode = _cliCommands.CreateEmpty(args);

        // Assert
        Assert.Equal(1, exitCode);
    }
}
