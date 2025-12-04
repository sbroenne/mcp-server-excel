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
        // Arrange & Act
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var result = _sheetCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Worksheets);
        Assert.NotEmpty(result.Worksheets); // Shared file has Sheet1 plus test sheets
    }
    /// <inheritdoc/>

    [Fact]
    public void Create_UniqueName_ReturnsSuccess()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"Create_{Guid.NewGuid():N}"[..31]; // Unique name, max 31 chars

        // Act
        _sheetCommands.Create(batch, sheetName);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var oldName = $"Old_{uniqueId}";
        var newName = $"New_{uniqueId}";
        _sheetCommands.Create(batch, oldName);

        // Act
        _sheetCommands.Rename(batch, oldName, newName);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = $"Del_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        // Act
        _sheetCommands.Delete(batch, sheetName);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var sourceName = $"Src_{uniqueId}";
        var targetName = $"Tgt_{uniqueId}";
        _sheetCommands.Create(batch, sourceName);

        // Act
        _sheetCommands.Copy(batch, sourceName, targetName);  // Copy throws on error

        // Assert - reaching here means copy succeeded

        // Verify both source and target sheets exist
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == sourceName);
        Assert.Contains(listResult.Worksheets, w => w.Name == targetName);
    }
}
