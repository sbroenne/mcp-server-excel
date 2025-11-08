using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// CLI-specific tests for DataModelCommands - verifying argument parsing, exit codes, and CLI behavior
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test argument validation (missing args, invalid args)
/// - ✅ Test exit code mapping (0 for success, 1 for error)
/// - ✅ Test user interaction (prompts, console output if applicable)
/// - ❌ DO NOT test Data Model operations or Excel COM interop (that's Core's responsibility)
///
/// These tests verify the CLI wrapper works correctly. Business logic is tested in ExcelMcp.Core.Tests.
/// </summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Feature", "DataModel")]
[Trait("Layer", "CLI")]
public class CliDataModelCommandsTests
{
    private readonly DataModelCommands _cliCommands;
    /// <inheritdoc/>

    public CliDataModelCommandsTests()
    {
        _cliCommands = new DataModelCommands();
    }
    /// <inheritdoc/>

    #region Argument Validation Tests

    [Fact]
    public void ListTables_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["dm-list-tables"]; // Missing file path

        // Act
        int exitCode = _cliCommands.ListTables(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void ListMeasures_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["dm-list-measures"]; // Missing file path

        // Act
        int exitCode = _cliCommands.ListMeasures(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Get_WithMissingMeasureNameArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["dm-view-measure", "SomeFile.xlsx"]; // Missing measure name

        // Act
        int exitCode = _cliCommands.ViewMeasure(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void ExportMeasure_WithMissingOutputFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["dm-export-measure", "SomeFile.xlsx", "SomeMeasure"]; // Missing output file

        // Act
        int exitCode = _cliCommands.ExportMeasure(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void ListRelationships_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["dm-list-relationships"]; // Missing file path

        // Act
        int exitCode = _cliCommands.ListRelationships(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void Refresh_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = ["dm-refresh"]; // Missing file path

        // Act
        int exitCode = _cliCommands.Refresh(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    #endregion

    #region Delete Operations Tests

    [Fact]
    public void DeleteMeasure_WithMissingArguments_ReturnsError()
    {
        // Arrange
        string[] args = ["dm-delete-measure", "SomeFile.xlsx"];

        // Act
        int exitCode = _cliCommands.DeleteMeasure(args);

        // Assert - CLI returns 1 for missing arguments
        Assert.Equal(1, exitCode);
    }
    /// <inheritdoc/>

    [Fact]
    public void DeleteRelationship_WithMissingArguments_ReturnsError()
    {
        // Arrange - Missing columns
        string[] args = ["dm-delete-relationship", "SomeFile.xlsx", "Table1", "Col1"];

        // Act
        int exitCode = _cliCommands.DeleteRelationship(args);

        // Assert - CLI returns 1 for missing arguments
        Assert.Equal(1, exitCode);
    }

    #endregion
}
