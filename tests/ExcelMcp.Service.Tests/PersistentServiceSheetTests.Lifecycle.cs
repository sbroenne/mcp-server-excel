using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for Sheet lifecycle operations (list, create, delete, rename, copy)
/// </summary>
public sealed partial class PersistentServiceSheetTests
{
    /// <inheritdoc/>
    [Fact]
    public void List_DefaultWorkbook_ReturnsDefaultSheets()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
        var result = _sheetCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Worksheets);
        Assert.NotEmpty(result.Worksheets); // Shared file has Sheet1 plus test sheets
    }

    /// <summary>
    /// Regression test: Visible property must be correctly read from Excel.
    /// Previously defaulted to false without reading sheet.Visible.
    /// </summary>
    [Fact]
    public void List_VisibleSheets_ReturnsVisibleTrue()
    {
        // Arrange & Act
        var batch = _fixture.BatchToken;
        var result = _sheetCommands.List(batch);

        // Assert - all default sheets should be visible
        Assert.True(result.Success);
        Assert.NotEmpty(result.Worksheets);
        Assert.All(result.Worksheets, sheet =>
            Assert.True(sheet.Visible, $"Sheet '{sheet.Name}' should be visible but Visible={sheet.Visible}"));
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_UniqueName_ReturnsSuccess()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"Create_{Guid.NewGuid():N}"[..31]; // Unique name, max 31 chars

        // Act
        _fixture.CreateNamedTestSheet(batch, sheetName);
        // Create throws on error, so reaching here means success

        // Verify sheet actually exists
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == sheetName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Rename_ExistingSheet_ReturnsSuccess()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var oldName = $"Old_{uniqueId}";
        var newName = $"New_{uniqueId}";
        _fixture.CreateNamedTestSheet(batch, oldName);

        // Act
        _sheetCommands.Rename(batch, oldName, newName);
        _fixture.RenameTrackedSheet(oldName, newName);
        // Rename throws on error, so reaching here means success

        // Verify rename actually happened
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == oldName);
        Assert.Contains(listResult.Worksheets, w => w.Name == newName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Delete_NonActiveSheet_ReturnsSuccess()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = $"Del_{Guid.NewGuid():N}"[..31];
        _fixture.CreateNamedTestSheet(batch, sheetName);

        // Act
        _sheetCommands.Delete(batch, sheetName);
        _fixture.ForgetSheet(sheetName);
        // Delete throws on error, so reaching here means success

        // Verify sheet is actually gone
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == sheetName);
    }
    /// <inheritdoc/>

    [Fact]
    public void Copy_ExistingSheet_CreatesNewSheet()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var sourceName = $"Src_{uniqueId}";
        var targetName = $"Tgt_{uniqueId}";
        _fixture.CreateNamedTestSheet(batch, sourceName);

        // Act
        _sheetCommands.Copy(batch, sourceName, targetName);  // Copy throws on error
        _fixture.RegisterSheetForCleanup(targetName);

        // Assert - reaching here means copy succeeded

        // Verify both source and target sheets exist
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == sourceName);
        Assert.Contains(listResult.Worksheets, w => w.Name == targetName);
    }
}



