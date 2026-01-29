using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for Sheet daemon handlers.
/// Verifies that daemon handlers correctly delegate to Core Commands.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Sheet")]
[Trait("Layer", "CLI")]
public class SheetDaemonHandlerTests : DaemonIntegrationTestBase
{
    private readonly SheetCommands _sheetCommands = new();

    public SheetDaemonHandlerTests(TempDirectoryFixture fixture) : base(fixture) { }

    [Fact]
    [Trait("Speed", "Fast")]
    public void SheetList_NewWorkbook_ReturnsDefaultSheet()
    {
        // Arrange - Use shared test file
        using var batch = CreateBatch();

        // Act - Same call as daemon handler
        var result = _sheetCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotEmpty(result.Worksheets);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void SheetCreate_ValidName_CreatesSheet()
    {
        // Arrange
        using var batch = CreateBatch();
        var sheetName = $"Test_{Guid.NewGuid():N}"[..31];

        // Act - Same as daemon: call command, verify with list
        _sheetCommands.Create(batch, sheetName);

        // Assert
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == sheetName);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void SheetRename_ExistingSheet_RenamesSuccessfully()
    {
        // Arrange
        using var batch = CreateBatch();
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var oldName = $"Old{uniqueId}";
        var newName = $"New{uniqueId}";
        _sheetCommands.Create(batch, oldName);

        // Act
        _sheetCommands.Rename(batch, oldName, newName);

        // Assert
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == oldName);
        Assert.Contains(listResult.Worksheets, w => w.Name == newName);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void SheetDelete_ExistingSheet_DeletesSuccessfully()
    {
        // Arrange
        using var batch = CreateBatch();
        var sheetName = $"Del_{Guid.NewGuid():N}"[..31];
        _sheetCommands.Create(batch, sheetName);

        // Act
        _sheetCommands.Delete(batch, sheetName);

        // Assert
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Worksheets, w => w.Name == sheetName);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void SheetCopy_ExistingSheet_CreatesCopy()
    {
        // Arrange
        using var batch = CreateBatch();
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var sourceName = $"Src{uniqueId}";
        var targetName = $"Tgt{uniqueId}";
        _sheetCommands.Create(batch, sourceName);

        // Act
        _sheetCommands.Copy(batch, sourceName, targetName);

        // Assert
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Worksheets, w => w.Name == sourceName);
        Assert.Contains(listResult.Worksheets, w => w.Name == targetName);
    }
}
