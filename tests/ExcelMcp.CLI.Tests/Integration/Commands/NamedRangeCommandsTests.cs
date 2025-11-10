using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// CLI-specific tests for NamedRangeCommands and CellCommands - verifying argument parsing, exit codes, and CLI behavior
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test argument validation (missing args, invalid args)
/// - ✅ Test exit code mapping (0 for success, 1 for error)
/// - ✅ Test user interaction (prompts, console output if applicable)
/// - ❌ DO NOT test named range or cell operations or Excel COM interop (that's Core's responsibility)
///
/// These tests verify the CLI wrapper works correctly. Business logic is tested in ExcelMcp.Core.Tests.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Parameters")]
[Trait("Layer", "CLI")]
public class CliNamedRangeCommandsTests
{
    private readonly NamedRangeCommands _cliCommands;
    /// <inheritdoc/>

    public CliNamedRangeCommandsTests()
    {
        _cliCommands = new NamedRangeCommands();
    }
    /// <inheritdoc/>

    [Fact]
    public void List_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["namedrange-list"]; // Missing file path

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_WithMissingParameterNameArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["namedrange-get", "file.xlsx"]; // Missing parameter name

        // Act
        int exitCode = _cliCommands.GetValue(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Set_WithMissingValueArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["namedrange-set", "file.xlsx", "ParamName"]; // Missing value

        // Act
        int exitCode = _cliCommands.SetValue(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_WithMissingReferenceArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["namedrange-create", "file.xlsx", "ParamName"]; // Missing reference

        // Act
        int exitCode = _cliCommands.Create(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Set_WithInvalidFileExtension_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["namedrange-set", "invalid.txt", "ParamName", "Value"];

        // Act
        int exitCode = _cliCommands.SetValue(args);

        // Assert - CLI returns 1 for error (invalid file extension)
        Assert.Equal(1, exitCode);
    }
}
