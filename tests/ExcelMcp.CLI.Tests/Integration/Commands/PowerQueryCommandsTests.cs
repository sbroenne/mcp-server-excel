using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// CLI-specific tests for PowerQueryCommands - verifying argument parsing, exit codes, and CLI behavior
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test argument validation (missing args, invalid args)
/// - ✅ Test exit code mapping (0 for success, 1 for error)
/// - ✅ Test user interaction (prompts, console output if applicable)
/// - ❌ DO NOT test Power Query M code operations or Excel COM interop (that's Core's responsibility)
///
/// These tests verify the CLI wrapper works correctly. Business logic is tested in ExcelMcp.Core.Tests.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Feature", "PowerQuery")]
[Trait("Layer", "CLI")]
public class CliPowerQueryCommandsTests
{
    private readonly PowerQueryCommands _cliCommands;
    /// <inheritdoc/>

    public CliPowerQueryCommandsTests()
    {
        _cliCommands = new PowerQueryCommands();
    }
    /// <inheritdoc/>

    [Fact]
    public void List_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["pq-list"]; // Missing file path

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void View_WithMissingArgs_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["pq-view", "file.xlsx"]; // Missing query name

        // Act
        int exitCode = _cliCommands.View(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Refresh_WithInvalidFileExtension_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["pq-refresh", "invalid.txt", "SomeQuery"];

        // Act
        int exitCode = _cliCommands.Refresh(args);

        // Assert - CLI returns 1 for error (invalid file extension)
        Assert.Equal(1, exitCode);
    }
}
