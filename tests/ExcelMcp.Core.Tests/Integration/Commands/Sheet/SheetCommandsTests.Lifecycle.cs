using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Tests for Sheet lifecycle operations (list, create, delete, rename, copy)
/// </summary>
public partial class SheetCommandsTests
{
    [Fact]
    public async Task List_DefaultWorkbook_ReturnsDefaultSheets()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(List_DefaultWorkbook_ReturnsDefaultSheets), _tempDir);

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _sheetCommands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Worksheets);
        Assert.NotEmpty(result.Worksheets); // New Excel file has Sheet1
    }

    [Fact]
    public async Task Create_UniqueName_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(Create_UniqueName_ReturnsSuccess), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Act
        var result = await _sheetCommands.CreateAsync(batch, "TestSheet");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");

        // Verify sheet actually exists
        var listResult = await _sheetCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == "TestSheet");

        // Save changes
    }

    [Fact]
    public async Task Rename_ExistingSheet_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(Rename_ExistingSheet_ReturnsSuccess), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "OldName");
        
        // Act
        var result = await _sheetCommands.RenameAsync(batch, "OldName", "NewName");

        // Assert
        Assert.True(result.Success, $"Rename failed: {result.ErrorMessage}");

        // Verify rename actually happened
        var listResult = await _sheetCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "OldName");
        Assert.Contains(listResult.Worksheets, w => w.Name == "NewName");

        // Save changes
    }

    [Fact]
    public async Task Delete_NonActiveSheet_ReturnsSuccess()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(Delete_NonActiveSheet_ReturnsSuccess), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "ToDelete");
        
        // Act
        var result = await _sheetCommands.DeleteAsync(batch, "ToDelete");

        // Assert
        Assert.True(result.Success, $"Delete failed: {result.ErrorMessage}");

        // Verify sheet is actually gone
        var listResult = await _sheetCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "ToDelete");

        // Save changes
    }

    [Fact]
    public async Task Copy_ExistingSheet_CreatesNewSheet()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SheetCommandsTests), nameof(Copy_ExistingSheet_CreatesNewSheet), _tempDir);

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        await _sheetCommands.CreateAsync(batch, "Source");
        
        // Act
        var result = await _sheetCommands.CopyAsync(batch, "Source", "Target");

        // Assert
        Assert.True(result.Success, $"Copy failed: {result.ErrorMessage}");

        // Verify both source and target sheets exist
        var listResult = await _sheetCommands.ListAsync(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == "Source");
        Assert.Contains(listResult.Worksheets, w => w.Name == "Target");

        // Save changes
    }
}
