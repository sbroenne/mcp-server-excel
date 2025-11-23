using Sbroenne.ExcelMcp.ComInterop.Session;
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
    public void List_DefaultWorkbook_ReturnsDefaultSheets()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(List_DefaultWorkbook_ReturnsDefaultSheets), _tempDir);

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _sheetCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Worksheets);
        Assert.NotEmpty(result.Worksheets); // New Excel file has Sheet1
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_UniqueName_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Create_UniqueName_ReturnsSuccess), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        _sheetCommands.Create(batch, "TestSheet");
        // Create throws on error, so reaching here means success

        // Verify sheet actually exists
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == "TestSheet");

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void Rename_ExistingSheet_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Rename_ExistingSheet_ReturnsSuccess), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "OldName");

        // Act
        _sheetCommands.Rename(batch, "OldName", "NewName");
        // Rename throws on error, so reaching here means success

        // Verify rename actually happened
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "OldName");
        Assert.Contains(listResult.Worksheets, w => w.Name == "NewName");

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void Delete_NonActiveSheet_ReturnsSuccess()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Delete_NonActiveSheet_ReturnsSuccess), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "ToDelete");

        // Act
        _sheetCommands.Delete(batch, "ToDelete");
        // Delete throws on error, so reaching here means success

        // Verify sheet is actually gone
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == "ToDelete");

        // Save changes
    }
    /// <inheritdoc/>

    [Fact]
    public void Copy_ExistingSheet_CreatesNewSheet()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Copy_ExistingSheet_CreatesNewSheet), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Source");

        // Act
        _sheetCommands.Copy(batch, "Source", "Target");  // Copy throws on error

        // Assert - reaching here means copy succeeded

        // Verify both source and target sheets exist
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == "Source");
        Assert.Contains(listResult.Worksheets, w => w.Name == "Target");

        // Save changes
    }
}
