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
    public void Begin_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["batch-begin"]; // Missing file path

        // Act
        int exitCode = _cliCommands.Begin(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Begin_WithNonExistentFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
        string[] args = ["batch-begin", nonExistentFile];

        // Act
        int exitCode = _cliCommands.Begin(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Commit_WithMissingBatchIdArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["batch-commit"]; // Missing batch ID

        // Act
        int exitCode = _cliCommands.Commit(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Commit_WithInvalidBatchId_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["batch-commit", "invalid-batch-id-12345"];

        // Act
        int exitCode = _cliCommands.Commit(args);

        // Assert - CLI returns 1 for error (batch not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void List_WithNoActiveBatches_ReturnsSuccessExitCode()
    {
        // Arrange
        string[] args = ["batch-list"];

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 0 even when no batches (success case)
        Assert.Equal(0, exitCode);
    }
}
