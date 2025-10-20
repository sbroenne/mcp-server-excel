using Xunit;
using Sbroenne.ExcelMcp.CLI.Commands;
using System.IO;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// Tests for CLI ParameterCommands - verifying CLI-specific behavior (argument parsing, exit codes)
/// These tests focus on the presentation layer, not the business logic
/// Core data logic is tested in ExcelMcp.Core.Tests
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Parameters")]
[Trait("Layer", "CLI")]
public class ParameterCommandsTests : IDisposable
{
    private readonly ParameterCommands _cliCommands;
    private readonly string _tempDir;

    public ParameterCommandsTests()
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

/// <summary>
/// Tests for CLI CellCommands - verifying CLI-specific behavior (argument parsing, exit codes)
/// These tests focus on the presentation layer, not the business logic
/// Core data logic is tested in ExcelMcp.Core.Tests
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Cells")]
[Trait("Layer", "CLI")]
public class CellCommandsTests : IDisposable
{
    private readonly CellCommands _cliCommands;
    private readonly string _tempDir;

    public CellCommandsTests()
    {
        _cliCommands = new CellCommands();
        
        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCLI_CellTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    [Fact]
    public void GetValue_WithMissingCellAddressArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "cell-get-value", "file.xlsx", "Sheet1" }; // Missing cell address

        // Act
        int exitCode = _cliCommands.GetValue(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void SetValue_WithMissingValueArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "cell-set-value", "file.xlsx", "Sheet1", "A1" }; // Missing value

        // Act
        int exitCode = _cliCommands.SetValue(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void GetFormula_WithMissingSheetNameArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "cell-get-formula", "file.xlsx" }; // Missing sheet name

        // Act
        int exitCode = _cliCommands.GetFormula(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void SetFormula_WithNonExistentFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentFile = Path.Combine(_tempDir, "NonExistent.xlsx");
        string[] args = { "cell-set-formula", nonExistentFile, "Sheet1", "A1", "=SUM(B1:B10)" };

        // Act
        int exitCode = _cliCommands.SetFormula(args);

        // Assert - CLI returns 1 for error (file not found)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void GetValue_WithInvalidFileExtension_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "cell-get-value", "invalid.txt", "Sheet1", "A1" };

        // Act
        int exitCode = _cliCommands.GetValue(args);

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