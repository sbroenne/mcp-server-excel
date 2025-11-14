using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// CLI-specific tests for BatchCommands - verifying argument parsing, exit codes, and CLI behavior
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test argument validation (missing args, invalid args)
/// - ✅ Test exit code mapping (0 for success, 1 for error)
/// - ❌ DO NOT test Excel COM interop (that's Core's responsibility)
///
/// These tests verify the CLI wrapper works correctly.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Batch")]
[Trait("Layer", "CLI")]
public class CliBatchCommandsTests
{
    private readonly BatchCommands _cliCommands;

    public CliBatchCommandsTests()
    {
        _cliCommands = new BatchCommands();
    }

    [Fact]
    public void Open_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["open"]; // Missing file path

        // Act
        int exitCode = _cliCommands.Open(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Open_WithNonExistentFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
        string[] args = ["open", nonExistentFile];

        // Act
        int exitCode = _cliCommands.Open(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Save_WithMissingSessionIdArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["save"]; // Missing session ID

        // Act
        int exitCode = _cliCommands.Save(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Save_WithInvalidSessionId_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["save", "invalid-session-id-12345"];

        // Act
        int exitCode = _cliCommands.Save(args);

        // Assert - CLI returns 1 for error (session not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void List_WithNoActiveSessions_ReturnsSuccessExitCode()
    {
        // Arrange
        string[] args = ["list"];

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 0 even when no sessions (success case)
        Assert.Equal(0, exitCode);
    }
}
