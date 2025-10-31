using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands.Others;

/// <summary>
/// Essential business logic tests for FileCommands - core functionality only
///
/// WHAT LLMs NEED TO KNOW:
/// 1. CreateEmpty creates .xlsx or .xlsm files (valid extensions)
/// 2. CreateEmpty fails on invalid extensions (.xls, .csv, .txt, etc.)
/// 3. CreateEmpty respects overwrite flag (default: fail if exists)
/// 4. TestFile validates existence and extension
/// 5. Result objects have Success, ErrorMessage, FilePath properties
///
/// REMOVED UNNECESSARY TESTS:
/// - Relative path conversion (Path.GetFullPath() handles this - not our concern)
/// - Nested directory creation (Directory.CreateDirectory() handles this - .NET's concern)
/// - Multiple file loops (just repeats basic test - no new knowledge)
/// - Invalid path characters (OS validation - not business logic)
/// - Timestamp verification (implementation detail - flaky tests)
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test Excel COM file operations and Result objects
/// - ✅ Test business rules (valid extensions, overwrite behavior)
/// - ❌ DO NOT test CLI argument parsing (CLI's responsibility)
/// - ❌ DO NOT test JSON serialization (MCP Server's responsibility)
/// - ❌ DO NOT test infrastructure (paths, directories, OS validation)
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Files")]
[Trait("Layer", "Core")]
public class FileCommandsTests : IDisposable
{
    private readonly FileCommands _fileCommands;
    private readonly string _tempDir;
    private readonly List<string> _createdFiles;

    public FileCommandsTests()
    {
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_FileTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _createdFiles = [];
    }

    /// <summary>
    /// Helper method to verify a file is a valid Excel workbook by trying to open it
    /// </summary>
    private async Task<bool> IsValidExcelFileAsync(string filePath)
    {
        try
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await batch.Execute<bool>((ctx, ct) =>
            {
                // If we can access the workbook and get worksheets, it's valid
                dynamic sheets = ctx.Book.Worksheets;
                return sheets.Count >= 1;
            });
        }
        catch
        {
            return false;
        }
    }

    [Fact]
    public async Task CreateEmpty_ValidXlsx_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "TestFile.xlsx");
        _createdFiles.Add(testFile);

        // Act
        var result = await _fileCommands.CreateEmptyAsync(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.Equal("create-empty", result.Action);
        Assert.NotNull(result.FilePath);
        Assert.True(File.Exists(testFile));

        // Verify it's a valid Excel workbook
        bool isValidExcel = await IsValidExcelFileAsync(testFile);
        Assert.True(isValidExcel, "Created file should be a valid Excel workbook with at least one worksheet");
    }

    [Fact]
    public async Task CreateEmpty_ValidXlsm_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "TestFile.xlsm");
        _createdFiles.Add(testFile);

        // Act
        var result = await _fileCommands.CreateEmptyAsync(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.True(File.Exists(testFile));

        // Verify it's a valid Excel workbook
        bool isValidExcel = await IsValidExcelFileAsync(testFile);
        Assert.True(isValidExcel, "Created file should be a valid Excel workbook");
    }

    [Theory]
    [InlineData("TestFile.xls")]
    [InlineData("TestFile.csv")]
    [InlineData("TestFile.txt")]
    public async Task CreateEmpty_InvalidExtension_ReturnsError(string fileName)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);

        // Act
        var result = await _fileCommands.CreateEmptyAsync(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("extension", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(testFile));
    }

    [Fact]
    public async Task CreateEmpty_FileExists_WithoutOverwrite_ReturnsError()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "ExistingFile.xlsx");
        _createdFiles.Add(testFile);

        // Create file first
        var firstResult = await _fileCommands.CreateEmptyAsync(testFile);
        Assert.True(firstResult.Success);

        // Act - Try to create again without overwrite flag
        var result = await _fileCommands.CreateEmptyAsync(testFile, overwriteIfExists: false);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task CreateEmpty_FileExists_WithOverwrite_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "OverwriteFile.xlsx");
        _createdFiles.Add(testFile);

        // Create file first
        var firstResult = await _fileCommands.CreateEmptyAsync(testFile);
        Assert.True(firstResult.Success);

        // Act - Overwrite
        var result = await _fileCommands.CreateEmptyAsync(testFile, overwriteIfExists: true);

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.True(File.Exists(testFile));
    }

    [Fact]
    public async Task TestFile_ExistingValidFile_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "TestValidation.xlsx");
        _createdFiles.Add(testFile);

        // Create a valid file
        File.WriteAllText(testFile, "dummy Excel content");

        // Act
        var result = await _fileCommands.TestFileAsync(testFile);

        // Assert
        Assert.True(result.Success);
        Assert.Null(result.ErrorMessage);
        Assert.True(result.Exists);
        Assert.True(result.IsValid);
        Assert.Equal(".xlsx", result.Extension);
        Assert.True(result.Size > 0);
    }

    [Fact]
    public async Task TestFile_NonExistent_ReturnsFailure()
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, "NonExistent.xlsx");

        // Act
        var result = await _fileCommands.TestFileAsync(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.False(result.Exists);
        Assert.False(result.IsValid);
    }

    [Theory]
    [InlineData("TestFile.xls", ".xls")]
    [InlineData("TestFile.csv", ".csv")]
    [InlineData("TestFile.txt", ".txt")]
    public async Task TestFile_InvalidExtension_ReturnsFailure(string fileName, string expectedExt)
    {
        // Arrange
        string testFile = Path.Combine(_tempDir, fileName);

        // Create file with invalid extension
        File.WriteAllText(testFile, "test content");
        _createdFiles.Add(testFile);

        // Act
        var result = await _fileCommands.TestFileAsync(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("Invalid file extension", result.ErrorMessage);
        Assert.True(result.Exists);
        Assert.False(result.IsValid);
        Assert.Equal(expectedExt, result.Extension);
    }


    public void Dispose()
    {
        // Clean up test files
        // NOTE: ExcelSession handles all COM cleanup synchronously, so files are already released
        // by the time tests complete. Just need basic file deletion.
        try
        {
            // Brief delay for any pending async operations
            System.Threading.Thread.Sleep(200);

            // Delete individual files first
            foreach (string file in _createdFiles)
            {
                try
                {
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                    // Best effort cleanup
                }
            }

            // Then delete the temp directory
            if (Directory.Exists(_tempDir))
            {
                try
                {
                    Directory.Delete(_tempDir, true);
                }
                catch
                {
                    // Best effort cleanup - don't fail tests if cleanup fails
                }
            }
        }
        catch
        {
            // Best effort cleanup - don't fail tests if cleanup fails
        }

        GC.SuppressFinalize(this);
    }
}
