using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Simple integration tests for FileCommands using batch pattern
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Files")]
[Trait("RequiresExcel", "true")]
public class FileCommandsSimpleTests : IDisposable
{
    private readonly string _testDir;
    private readonly FileCommands _commands;

    public FileCommandsSimpleTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_FileSimple_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
        _commands = new FileCommands();
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, recursive: true);
            }
        }
        catch { /* Cleanup failure is non-critical */ }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task CreateEmpty_ValidPath_CreatesFile()
    {
        // Arrange
        var testFile = Path.Combine(_testDir, "test.xlsx");

        // Act
        var result = await _commands.CreateEmptyAsync(testFile, overwriteIfExists: false);

        // Assert
        Assert.True(result.Success, $"CreateEmpty failed: {result.ErrorMessage}");
        Assert.True(File.Exists(testFile), "File was not created");
        Assert.Equal(testFile, result.FilePath);
    }

    [Fact]
    public async Task CreateEmpty_XlsmExtension_CreatesFile()
    {
        // Arrange
        var testFile = Path.Combine(_testDir, "test.xlsm");

        // Act
        var result = await _commands.CreateEmptyAsync(testFile, overwriteIfExists: false);

        // Assert
        Assert.True(result.Success, $"CreateEmpty failed: {result.ErrorMessage}");
        Assert.True(File.Exists(testFile), "File was not created");
        Assert.Equal(testFile, result.FilePath);
    }

    [Fact]
    public async Task CreateEmpty_ExistingFileNoOverwrite_ReturnsError()
    {
        // Arrange
        var testFile = Path.Combine(_testDir, "test.xlsx");
        await _commands.CreateEmptyAsync(testFile, overwriteIfExists: false);

        // Act
        var result = await _commands.CreateEmptyAsync(testFile, overwriteIfExists: false);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }
}
