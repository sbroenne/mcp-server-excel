using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.File;

/// <summary>
/// Tests for FileCommands CreateEmpty operation
/// </summary>
public partial class FileCommandsTests
{
    /// <summary>
    /// Helper method to verify a file is a valid Excel workbook by trying to open it
    /// </summary>
    private static Task<bool> IsValidExcelFileAsync(string filePath)
    {
        try
        {
            using var batch = ExcelSession.BeginBatch(filePath);
            var isValid = batch.Execute((ctx, ct) =>
            {
                // If we can access the workbook and get worksheets, it's valid
                dynamic sheets = ctx.Book.Worksheets;
                return sheets.Count >= 1;
            });
            return Task.FromResult(isValid);
        }
        catch (Exception)
        {
            // Test helper - any Excel error means file is invalid
            return Task.FromResult(false);
        }
    }
    /// <inheritdoc/>

    [Fact]
    public async Task CreateEmpty_ValidXlsx_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{nameof(CreateEmpty_ValidXlsx_ReturnsSuccess)}_{Guid.NewGuid():N}.xlsx");

        // Act
        var result = await _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Null(result.ErrorMessage);
        Assert.Equal("create-empty", result.Action);
        Assert.NotNull(result.FilePath);
        Assert.True(System.IO.File.Exists(testFile));

        // Verify it's a valid Excel workbook
        bool isValidExcel = await IsValidExcelFileAsync(testFile);
        Assert.True(isValidExcel, "Created file should be a valid Excel workbook with at least one worksheet");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task CreateEmpty_ValidXlsm_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{nameof(CreateEmpty_ValidXlsm_ReturnsSuccess)}_{Guid.NewGuid():N}.xlsm");

        // Act
        var result = await _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Null(result.ErrorMessage);
        Assert.True(System.IO.File.Exists(testFile));

        // Verify it's a valid Excel workbook
        bool isValidExcel = await IsValidExcelFileAsync(testFile);
        Assert.True(isValidExcel, "Created file should be a valid Excel workbook");
    }
    /// <inheritdoc/>

    [Theory]
    [InlineData("TestFile.xls")]
    [InlineData("TestFile.csv")]
    [InlineData("TestFile.txt")]
    public async Task CreateEmpty_InvalidExtension_ReturnsError(string fileName)
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{Guid.NewGuid():N}_{fileName}");

        // Act
        var result = await _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("extension", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.False(System.IO.File.Exists(testFile));
    }
    /// <inheritdoc/>

    [Fact]
    public async Task CreateEmpty_FileExists_WithoutOverwrite_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(FileCommandsTests), nameof(CreateEmpty_FileExists_WithoutOverwrite_ReturnsError), _tempDir);

        // Act - Try to create again without overwrite flag
        var result = await _fileCommands.CreateEmpty(testFile, overwriteIfExists: false);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("already exists", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public async Task CreateEmpty_FileExists_WithOverwrite_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(FileCommandsTests), nameof(CreateEmpty_FileExists_WithOverwrite_ReturnsSuccess), _tempDir);

        // Act - Overwrite
        var result = await _fileCommands.CreateEmpty(testFile, overwriteIfExists: true);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Null(result.ErrorMessage);
        Assert.True(System.IO.File.Exists(testFile));
    }
}
