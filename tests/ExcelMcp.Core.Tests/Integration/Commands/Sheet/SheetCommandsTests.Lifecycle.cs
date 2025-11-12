using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Tests for Sheet lifecycle operations (list, create, delete, rename, copy)
/// </summary>
public partial class SheetCommandsTests
{
    /// <inheritdoc/>
    [Fact]
    public async Task List_DefaultWorkbook_ReturnsDefaultSheets()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(List_DefaultWorkbook_ReturnsDefaultSheets), _tempDir);

        // Act - Use filePath-based API
        var result = await _sheetCommands.ListAsync(testFile);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Worksheets);
        Assert.NotEmpty(result.Worksheets); // New Excel file has Sheet1
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Create_UniqueName_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(Create_UniqueName_ReturnsSuccess), _tempDir);

        // Act - Use filePath-based API
        var result = await _sheetCommands.CreateAsync(testFile, "TestSheet");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");

        // Verify sheet actually exists
        var listResult = await _sheetCommands.ListAsync(testFile);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == "TestSheet");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Rename_ExistingSheet_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(Rename_ExistingSheet_ReturnsSuccess), _tempDir);

        await _sheetCommands.CreateAsync(testFile, "OldName");

        // Act - Use filePath-based API
        var result = await _sheetCommands.RenameAsync(testFile, "OldName", "NewName");

        // Assert
        Assert.True(result.Success, $"Rename failed: {result.ErrorMessage}");

        // Verify rename actually happened
        var listResult = await _sheetCommands.ListAsync(testFile);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "OldName");
        Assert.Contains(listResult.Worksheets, w => w.Name == "NewName");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Delete_NonActiveSheet_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(Delete_NonActiveSheet_ReturnsSuccess), _tempDir);

        await _sheetCommands.CreateAsync(testFile, "ToDelete");

        // Act - Use filePath-based API
        var result = await _sheetCommands.DeleteAsync(testFile, "ToDelete");

        // Assert
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Verify sheet is actually gone
        var listResult = await _sheetCommands.ListAsync(testFile);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "ToDelete");
    }
    /// <inheritdoc/>

    [Fact]
    public async Task Copy_ExistingSheet_CreatesNewSheet()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(Copy_ExistingSheet_CreatesNewSheet), _tempDir);

        await _sheetCommands.CreateAsync(testFile, "Source");

        // Act - Use filePath-based API
        var result = await _sheetCommands.CopyAsync(testFile, "Source", "Target");

        // Assert
        Assert.True(result.Success, $"Copy failed: {result.ErrorMessage}");

        // Verify both source and target sheets exist
        var listResult = await _sheetCommands.ListAsync(testFile);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == "Source");
        Assert.Contains(listResult.Worksheets, w => w.Name == "Target");
    }
}
