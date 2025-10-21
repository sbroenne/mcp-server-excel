using Sbroenne.ExcelMcp.CLI.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration.Commands;

/// <summary>
/// Tests for CLI SheetCommands - verifying CLI-specific behavior (argument parsing, exit codes)
/// These tests focus on the presentation layer, not the business logic
/// Core data logic is tested in ExcelMcp.Core.Tests
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

        _createdFiles = new List<string>();
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
    public void Read_WithMissingArgs_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "sheet-read", "file.xlsx" }; // Missing sheet name and range

        // Act
        int exitCode = _cliCommands.Read(args);

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

    [Fact]
    public void Clear_WithMissingRangeArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "sheet-clear", "file.xlsx", "Sheet1" }; // Missing range

        // Act
        int exitCode = _cliCommands.Clear(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public async Task Write_WithMissingDataFileArg_ReturnsErrorExitCode()
    {
        // Arrange
        string[] args = { "sheet-write", "file.xlsx", "Sheet1" }; // Missing data file

        // Act
        int exitCode = await _cliCommands.Write(args);

        // Assert - CLI returns 1 for error (missing arguments)
        Assert.Equal(1, exitCode);
    }

    [Fact]
    public void Append_WithNonExistentDataFile_ReturnsErrorExitCode()
    {
        // Arrange
        string nonExistentDataFile = Path.Combine(_tempDir, "NonExistent.csv");
        string[] args = { "sheet-append", "file.xlsx", "Sheet1", nonExistentDataFile };

        // Act
        int exitCode = _cliCommands.Append(args);

        // Assert - CLI returns 1 for error (data file not found)
        Assert.Equal(1, exitCode);
    }

    public void Dispose()
    {
        // Clean up test files
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
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, true);
                        break;
                    }
                    catch (IOException)
                    {
                        if (i == 2) throw;
                        System.Threading.Thread.Sleep(1000);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
        }
        catch { }

        GC.SuppressFinalize(this);
    }
}