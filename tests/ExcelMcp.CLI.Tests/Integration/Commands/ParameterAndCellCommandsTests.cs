using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// CLI-specific tests for ParameterCommands and CellCommands - verifying argument parsing, exit codes, and CLI behavior
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test argument validation (missing args, invalid args)
/// - ✅ Test exit code mapping (0 for success, 1 for error)
/// - ✅ Test user interaction (prompts, console output if applicable)
/// - ❌ DO NOT test named range or cell operations or Excel COM interop (that's Core's responsibility)
///
/// These tests verify the CLI wrapper works correctly. Business logic is tested in ExcelMcp.Core.Tests.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Parameters")]
[Trait("Layer", "CLI")]
public class CliParameterCommandsTests : IDisposable
{
    private readonly ParameterCommands _cliCommands;
    private readonly string _tempDir;

    public CliParameterCommandsTests()
    {
        _cliCommands = new ParameterCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_ParameterTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public void List_WithMissingFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "param-list" }; // Missing file path

        // Act
        int exitCode = _cliCommands.List(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Get_WithMissingParameterNameArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "param-get", "file.xlsx" }; // Missing parameter name

        // Act
        int exitCode = _cliCommands.Get(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Set_WithMissingValueArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "param-set", "file.xlsx", "ParamName" }; // Missing value

        // Act
        int exitCode = _cliCommands.Set(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Create_WithMissingReferenceArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "param-create", "file.xlsx", "ParamName" }; // Missing reference

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
        string[] args = { "param-delete", nonExistentFile, "SomeParam" };

        // Act
        int exitCode = _cliCommands.Delete(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Set_WithInvalidFileExtension_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "param-set", "invalid.txt", "ParamName", "Value" };

        // Act
        int exitCode = _cliCommands.Set(args);

        // Assert - CLI returns 1 for error (invalid file extension)
        Assert.Equal(1, exitCode);
    }

    public void Dispose()
    {
        // Clean up temp directory
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch { }

        GC.SuppressFinalize(this);
    }
}
