using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Tests for Sheet move and cross-workbook operations (move, copy-to-workbook, move-to-workbook)
/// </summary>
public partial class SheetCommandsTests
{
    // ========================================
    // MOVE (within workbook) Tests
    // ========================================

    /// <inheritdoc/>
    [Fact]
    public void Move_WithBeforeSheet_RepositionsSheet()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_WithBeforeSheet_RepositionsSheet), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "MoveMe");
        _sheetCommands.Create(batch, "Target");

        // Get initial position of MoveMe
        _sheetCommands.List(batch);

        // Act - Move MoveMe before Sheet1
        var result = _sheetCommands.Move(batch, "MoveMe", beforeSheet: "Sheet1");

        // Assert
        Assert.True(result.Success, $"Move failed: {result.ErrorMessage}");

        // Verify MoveMe moved to a different position (should now be before Sheet1)
        var afterList = _sheetCommands.List(batch);
        var sheets = afterList.Worksheets.ToList();
        var afterIndex = sheets.FindIndex(s => s.Name == "MoveMe");
        var sheet1Index = sheets.FindIndex(s => s.Name == "Sheet1");

        // MoveMe should be before Sheet1
        Assert.True(afterIndex < sheet1Index, $"Expected MoveMe (index {afterIndex}) to be before Sheet1 (index {sheet1Index})");
    }

    /// <inheritdoc/>
    [Fact]
    public void Move_WithAfterSheet_RepositionsSheet()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_WithAfterSheet_RepositionsSheet), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "MoveMe");
        _sheetCommands.Create(batch, "Target");

        // Get initial position
        _sheetCommands.List(batch);

        // Act - Move MoveMe after Target
        var result = _sheetCommands.Move(batch, "MoveMe", afterSheet: "Target");

        // Assert
        Assert.True(result.Success, $"Move failed: {result.ErrorMessage}");

        // Verify MoveMe is now after Target
        var afterList = _sheetCommands.List(batch);
        var sheets = afterList.Worksheets.ToList();
        var afterIndex = sheets.FindIndex(s => s.Name == "MoveMe");
        var targetIndex = sheets.FindIndex(s => s.Name == "Target");

        // MoveMe should be after Target
        Assert.True(afterIndex > targetIndex, $"Expected MoveMe (index {afterIndex}) to be after Target (index {targetIndex})");
    }

    /// <inheritdoc/>
    [Fact]
    public void Move_NoPositionSpecified_MovesToEnd()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_NoPositionSpecified_MovesToEnd), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Sheet2");
        _sheetCommands.Create(batch, "Sheet3");

        // Act - Move Sheet1 without specifying position
        var result = _sheetCommands.Move(batch, "Sheet1");

        // Assert
        Assert.True(result.Success, $"Move failed: {result.ErrorMessage}");

        // Verify Sheet1 is now at the end
        var listResult = _sheetCommands.List(batch);
        Assert.True(listResult.Success);
        var sheets = listResult.Worksheets.ToList();
        Assert.Equal("Sheet1", sheets[^1].Name); // Last sheet
    }

    /// <inheritdoc/>
    [Fact]
    public void Move_BothBeforeAndAfter_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_BothBeforeAndAfter_ThrowsException), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Sheet2");
        _sheetCommands.Create(batch, "Sheet3");

        // Act & Assert - Should throw when both beforeSheet and afterSheet are specified
        var exception = Assert.Throws<ArgumentException>(
            () => _sheetCommands.Move(batch, "Sheet1", beforeSheet: "Sheet2", afterSheet: "Sheet3"));
        Assert.Contains("both beforeSheet and afterSheet", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void Move_NonExistentSheet_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_NonExistentSheet_ThrowsException), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert - Should throw when sheet doesn't exist
        var exception = Assert.Throws<InvalidOperationException>(
            () => _sheetCommands.Move(batch, "NonExistent", afterSheet: "Sheet1"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void Move_NonExistentTargetSheet_ThrowsException()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_NonExistentTargetSheet_ThrowsException), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Sheet2");

        // Act & Assert - Should throw when target sheet doesn't exist
        var exception = Assert.Throws<InvalidOperationException>(
            () => _sheetCommands.Move(batch, "Sheet2", beforeSheet: "NonExistent"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ========================================
    // COPY-TO-WORKBOOK (cross-workbook) Tests
    // Tests for copying sheets between different workbooks using multi-file batch
    // ========================================

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_WithTargetName_CopiesAndRenames()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithTargetName_CopiesAndRenames)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithTargetName_CopiesAndRenames)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "SourceSheet", sourceFile);

        // Act - Copy sheet to target workbook with new name
        var result = _sheetCommands.CopyToWorkbook(batch, sourceFile, "SourceSheet", targetFile, "CopiedSheet");

        // Assert
        Assert.True(result.Success, $"CopyToWorkbook failed: {result.ErrorMessage}");

        // Verify sheet exists in target workbook with new name
        var targetList = _sheetCommands.List(batch, targetFile);
        Assert.Contains(targetList.Worksheets, s => s.Name == "CopiedSheet");

        // Verify source sheet still exists in source workbook
        var sourceList = _sheetCommands.List(batch, sourceFile);
        Assert.Contains(sourceList.Worksheets, s => s.Name == "SourceSheet");
    }

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_NoTargetName_CopiesWithOriginalName()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_NoTargetName_CopiesWithOriginalName)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_NoTargetName_CopiesWithOriginalName)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "SourceSheet", sourceFile);

        // Act - Copy without specifying target name
        var result = _sheetCommands.CopyToWorkbook(batch, sourceFile, "SourceSheet", targetFile);

        // Assert
        Assert.True(result.Success, $"CopyToWorkbook failed: {result.ErrorMessage}");

        // Verify sheet was copied (Excel keeps original name)
        var targetList = _sheetCommands.List(batch, targetFile);
        Assert.Contains(targetList.Worksheets, s => s.Name == "SourceSheet");
    }

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_WithBeforeSheet_PositionsCorrectly()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithBeforeSheet_PositionsCorrectly)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithBeforeSheet_PositionsCorrectly)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "SourceSheet", sourceFile);

        // Act - Copy before Sheet1 in target workbook
        var result = _sheetCommands.CopyToWorkbook(batch, sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1");

        // Assert
        Assert.True(result.Success, $"CopyToWorkbook failed: {result.ErrorMessage}");

        // Verify sheet was copied to target workbook
        var targetList = _sheetCommands.List(batch, targetFile);
        Assert.Contains(targetList.Worksheets, s => s.Name == "Copied");
    }

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_WithAfterSheet_PositionsCorrectly()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithAfterSheet_PositionsCorrectly)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithAfterSheet_PositionsCorrectly)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "SourceSheet", sourceFile);

        // Act - Copy after Sheet1 in target workbook
        var result = _sheetCommands.CopyToWorkbook(batch, sourceFile, "SourceSheet", targetFile, "Copied", afterSheet: "Sheet1");

        // Assert
        Assert.True(result.Success, $"CopyToWorkbook failed: {result.ErrorMessage}");

        // Verify sheet was copied to target workbook
        var targetList = _sheetCommands.List(batch, targetFile);
        Assert.Contains(targetList.Worksheets, s => s.Name == "Copied");
    }

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_BothBeforeAndAfter_ThrowsException()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_BothBeforeAndAfter_ThrowsException)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_BothBeforeAndAfter_ThrowsException)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "SourceSheet", sourceFile);

        // Act & Assert - Should throw ArgumentException when both beforeSheet and afterSheet specified
        var exception = Assert.Throws<ArgumentException>(
            () => _sheetCommands.CopyToWorkbook(batch, sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1", afterSheet: "Sheet1"));
        Assert.Contains("both beforeSheet and afterSheet", exception.Message, StringComparison.OrdinalIgnoreCase);
    }


    // ========================================
    // MOVE-TO-WORKBOOK (cross-workbook) Tests
    // Tests for moving sheets between different workbooks using multi-file batch
    // ========================================

    /// <inheritdoc/>
    [Fact]
    public void MoveToWorkbook_Default_MovesSheetSuccessfully()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_Default_MovesSheetSuccessfully)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_Default_MovesSheetSuccessfully)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "MoveMe", sourceFile);

        // Act - Move sheet to target workbook
        var result = _sheetCommands.MoveToWorkbook(batch, sourceFile, "MoveMe", targetFile);

        // Assert
        Assert.True(result.Success, $"MoveToWorkbook failed: {result.ErrorMessage}");

        // Verify sheet exists in target workbook
        var targetList = _sheetCommands.List(batch, targetFile);
        Assert.Contains(targetList.Worksheets, s => s.Name == "MoveMe");

        // Verify sheet no longer exists in source workbook
        var sourceList = _sheetCommands.List(batch, sourceFile);
        Assert.DoesNotContain(sourceList.Worksheets, s => s.Name == "MoveMe");
    }

    /// <inheritdoc/>
    [Fact]
    public void MoveToWorkbook_WithBeforeSheet_PositionsCorrectly()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_WithBeforeSheet_PositionsCorrectly)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_WithBeforeSheet_PositionsCorrectly)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "MoveMe", sourceFile);

        // Act - Move before Sheet1 in target workbook
        var result = _sheetCommands.MoveToWorkbook(batch, sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1");

        // Assert
        Assert.True(result.Success, $"MoveToWorkbook failed: {result.ErrorMessage}");

        // Verify sheet was moved to target workbook
        var targetList = _sheetCommands.List(batch, targetFile);
        Assert.Contains(targetList.Worksheets, s => s.Name == "MoveMe");

        // Verify removal from source
        var sourceList = _sheetCommands.List(batch, sourceFile);
        Assert.DoesNotContain(sourceList.Worksheets, s => s.Name == "MoveMe");
    }

    /// <inheritdoc/>
    [Fact]
    public void MoveToWorkbook_WithAfterSheet_PositionsCorrectly()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_WithAfterSheet_PositionsCorrectly)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_WithAfterSheet_PositionsCorrectly)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "MoveMe", sourceFile);

        // Act - Move after Sheet1 in target workbook
        var result = _sheetCommands.MoveToWorkbook(batch, sourceFile, "MoveMe", targetFile, afterSheet: "Sheet1");

        // Assert
        Assert.True(result.Success, $"MoveToWorkbook failed: {result.ErrorMessage}");

        // Verify sheet was moved to target workbook
        var targetList = _sheetCommands.List(batch, targetFile);
        Assert.Contains(targetList.Worksheets, s => s.Name == "MoveMe");

        // Verify removal from source
        var sourceList = _sheetCommands.List(batch, sourceFile);
        Assert.DoesNotContain(sourceList.Worksheets, s => s.Name == "MoveMe");
    }

    /// <inheritdoc/>
    [Fact]
    public void MoveToWorkbook_BothBeforeAndAfter_ThrowsException()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_BothBeforeAndAfter_ThrowsException)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_BothBeforeAndAfter_ThrowsException)}_Target", _tempDir);

        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);

        _sheetCommands.Create(batch, "MoveMe", sourceFile);

        // Act & Assert - Should throw ArgumentException when both beforeSheet and afterSheet specified
        var exception = Assert.Throws<ArgumentException>(
            () => _sheetCommands.MoveToWorkbook(batch, sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1", afterSheet: "Sheet1"));
        Assert.Contains("both beforeSheet and afterSheet", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
