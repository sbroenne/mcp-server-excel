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
    private static bool IsValidExcelFile(string filePath)
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
            return isValid;
        }
        catch (Exception)
        {
            // Test helper - any Excel error means file is invalid
            return false;
        }
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateEmpty_ValidXlsx_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{nameof(CreateEmpty_ValidXlsx_ReturnsSuccess)}_{Guid.NewGuid():N}.xlsx");

        // Act
        _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.True(System.IO.File.Exists(testFile));

        // Verify it's a valid Excel workbook
        bool isValidExcel = IsValidExcelFile(testFile);
        Assert.True(isValidExcel, "Created file should be a valid Excel workbook with at least one worksheet");
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateEmpty_ValidXlsm_ReturnsSuccess()
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{nameof(CreateEmpty_ValidXlsm_ReturnsSuccess)}_{Guid.NewGuid():N}.xlsm");

        // Act
        _fileCommands.CreateEmpty(testFile);

        // Assert
        Assert.True(System.IO.File.Exists(testFile));

        // Verify it's a valid Excel workbook
        bool isValidExcel = IsValidExcelFile(testFile);
        Assert.True(isValidExcel, "Created file should be a valid Excel workbook");
    }
    /// <inheritdoc/>

    [Theory]
    [InlineData("TestFile.xls")]
    [InlineData("TestFile.csv")]
    [InlineData("TestFile.txt")]
    public void CreateEmpty_InvalidExtension_ReturnsError(string fileName)
    {
        // Arrange
        string testFile = Path.Join(_tempDir, $"{Guid.NewGuid():N}_{fileName}");

        // Act
        var exception = Assert.Throws<ArgumentException>(() => _fileCommands.CreateEmpty(testFile));

        // Assert
        Assert.Contains("extension", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(System.IO.File.Exists(testFile));
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateEmpty_FileExists_WithoutOverwrite_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(FileCommandsTests), nameof(CreateEmpty_FileExists_WithoutOverwrite_ReturnsError), _tempDir);

        // Act - Try to create again without overwrite flag
        var exception = Assert.Throws<InvalidOperationException>(() =>
            _fileCommands.CreateEmpty(testFile, overwriteIfExists: false));

        // Assert
        Assert.Contains("already exists", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateEmpty_FileExists_WithOverwrite_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(FileCommandsTests), nameof(CreateEmpty_FileExists_WithOverwrite_ReturnsSuccess), _tempDir);

        // Act - Overwrite
        _fileCommands.CreateEmpty(testFile, overwriteIfExists: true);

        // Assert
        Assert.True(System.IO.File.Exists(testFile));
    }
}
