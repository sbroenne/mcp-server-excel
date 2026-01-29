using Sbroenne.ExcelMcp.CLI.Infrastructure;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "CLI")]
[Trait("Category", "Unit")]
[Trait("Feature", "PathValidation")]
[Trait("Speed", "Fast")]
public sealed class PathValidatorTests
{
    [Theory]
    [InlineData("test.xlsx", true)]
    [InlineData("test.xlsm", true)]
    [InlineData("test.xlsb", true)]
    [InlineData("test.xls", true)]
    [InlineData("test.XLSX", true)]
    [InlineData("test.csv", false)]
    [InlineData("test.txt", false)]
    [InlineData("test", false)]
    public void ValidateExcelPath_ValidatesExtension(string fileName, bool expectedValid)
    {
        var filePath = Path.Combine(Path.GetTempPath(), fileName);
        var result = PathValidator.ValidateExcelPath(filePath, mustExist: false, out var errorMessage);

        Assert.Equal(expectedValid, result);
        if (!expectedValid)
        {
            Assert.Contains("extension", errorMessage ?? "", StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void ValidateExcelPath_NullPath_ReturnsFalse()
    {
        var result = PathValidator.ValidateExcelPath(null, mustExist: false, out var errorMessage);

        Assert.False(result);
        Assert.Contains("required", errorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidateExcelPath_EmptyPath_ReturnsFalse()
    {
        var result = PathValidator.ValidateExcelPath("", mustExist: false, out var errorMessage);

        Assert.False(result);
        Assert.Contains("required", errorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ValidateExcelPath_MustExist_FileNotFound()
    {
        var filePath = Path.Combine(Path.GetTempPath(), "nonexistent_file_" + Guid.NewGuid() + ".xlsx");
        var result = PathValidator.ValidateExcelPath(filePath, mustExist: true, out var errorMessage);

        Assert.False(result);
        Assert.Contains("not found", errorMessage ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void NormalizePath_RelativePath_ReturnsFullPath()
    {
        var relativePath = "test.xlsx";
        var result = PathValidator.NormalizePath(relativePath);

        Assert.True(Path.IsPathRooted(result));
        Assert.EndsWith("test.xlsx", result);
    }

    [Fact]
    public void NormalizePath_AlreadyFullPath_ReturnsSame()
    {
        var fullPath = Path.Combine(Path.GetTempPath(), "test.xlsx");
        var result = PathValidator.NormalizePath(fullPath);

        Assert.Equal(fullPath, result);
    }

    [Fact]
    public void GetUniqueFilePath_FileDoesNotExist_ReturnsSamePath()
    {
        var filePath = Path.Combine(Path.GetTempPath(), "unique_" + Guid.NewGuid() + ".xlsx");
        var result = PathValidator.GetUniqueFilePath(filePath);

        Assert.Equal(filePath, result);
    }

    [Fact]
    public void GetUniqueFilePath_FileExists_ReturnsNumberedPath()
    {
        // Create a temp file
        var filePath = Path.Combine(Path.GetTempPath(), "exists_test_" + Guid.NewGuid() + ".xlsx");
        try
        {
            File.WriteAllText(filePath, "test");

            var result = PathValidator.GetUniqueFilePath(filePath);

            Assert.NotEqual(filePath, result);
            Assert.Contains("_1", result);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void HasInvalidPathCharacters_ValidPath_ReturnsFalse()
    {
        var validPath = @"C:\Users\test\Documents\file.xlsx";
        var result = PathValidator.HasInvalidPathCharacters(validPath);

        Assert.False(result);
    }

    [Fact]
    public void IsWithinAllowedDirectory_InsideDirectory_ReturnsTrue()
    {
        var allowedDir = Path.GetTempPath();
        var filePath = Path.Combine(allowedDir, "subdir", "file.xlsx");

        var result = PathValidator.IsWithinAllowedDirectory(filePath, allowedDir);

        Assert.True(result);
    }

    [Fact]
    public void IsWithinAllowedDirectory_OutsideDirectory_ReturnsFalse()
    {
        var allowedDir = Path.Combine(Path.GetTempPath(), "allowed_" + Guid.NewGuid());
        var filePath = Path.Combine(Path.GetTempPath(), "other_dir", "file.xlsx");

        var result = PathValidator.IsWithinAllowedDirectory(filePath, allowedDir);

        Assert.False(result);
    }
}
