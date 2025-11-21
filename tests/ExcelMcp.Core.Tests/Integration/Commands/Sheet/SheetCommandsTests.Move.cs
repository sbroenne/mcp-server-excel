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
    public void Move_BothBeforeAndAfter_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_BothBeforeAndAfter_ReturnsError), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Sheet2");
        _sheetCommands.Create(batch, "Sheet3");

        // Act - Try to specify both beforeSheet and afterSheet
        var result = _sheetCommands.Move(batch, "Sheet1", beforeSheet: "Sheet2", afterSheet: "Sheet3");

        // Assert
        Assert.False(result.Success, "Expected failure when both beforeSheet and afterSheet are specified");
        Assert.Contains("both beforeSheet and afterSheet", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void Move_NonExistentSheet_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_NonExistentSheet_ReturnsError), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Try to move a sheet that doesn't exist
        var result = _sheetCommands.Move(batch, "NonExistent", afterSheet: "Sheet1");

        // Assert
        Assert.False(result.Success, "Expected failure when sheet doesn't exist");
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void Move_NonExistentTargetSheet_ReturnsError()
    {
        // Arrange
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), nameof(Move_NonExistentTargetSheet_ReturnsError), _tempDir);

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Sheet2");

        // Act - Try to move relative to a sheet that doesn't exist
        var result = _sheetCommands.Move(batch, "Sheet2", beforeSheet: "NonExistent");

        // Assert
        Assert.False(result.Success, "Expected failure when target sheet doesn't exist");
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    // ========================================
    // COPY-TO-WORKBOOK (cross-workbook) Tests
    // NOTE: Cross-workbook operations not currently supported due to COM RCW limitations
    // These tests verify graceful error handling
    // ========================================

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_WithTargetName_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithTargetName_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithTargetName_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "SourceSheet");

        // Act - Attempt cross-workbook copy
        var result = _sheetCommands.CopyToWorkbook(sourceBatch, "SourceSheet", targetBatch, "CopiedSheet");

        // Assert - Should return clear error about limitation
        Assert.False(result.Success, "Expected failure for unsupported cross-workbook operation");
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("COM interop", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_NoTargetName_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_NoTargetName_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_NoTargetName_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "SourceSheet");

        // Act - Attempt cross-workbook copy
        var result = _sheetCommands.CopyToWorkbook(sourceBatch, "SourceSheet", targetBatch);

        // Assert - Should return clear error about limitation
        Assert.False(result.Success);
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_WithBeforeSheet_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithBeforeSheet_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithBeforeSheet_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "SourceSheet");
        _sheetCommands.Create(targetBatch, "Target2");

        // Act - Attempt cross-workbook copy
        var result = _sheetCommands.CopyToWorkbook(sourceBatch, "SourceSheet", targetBatch, "Copied", beforeSheet: "Sheet1");

        // Assert - Should return clear error
        Assert.False(result.Success);
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_WithAfterSheet_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithAfterSheet_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_WithAfterSheet_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "SourceSheet");
        _sheetCommands.Create(targetBatch, "Target2");

        // Act - Attempt cross-workbook copy
        var result = _sheetCommands.CopyToWorkbook(sourceBatch, "SourceSheet", targetBatch, "Copied", afterSheet: "Target2");

        // Assert - Should return clear error
        Assert.False(result.Success);
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void CopyToWorkbook_BothBeforeAndAfter_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_BothBeforeAndAfter_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(CopyToWorkbook_BothBeforeAndAfter_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "SourceSheet");

        // Act - Cross-workbook operation (parameter validation happens after "not supported" check)
        var result = _sheetCommands.CopyToWorkbook(sourceBatch, "SourceSheet", targetBatch, "Copied", beforeSheet: "Sheet1", afterSheet: "Sheet1");

        // Assert - Should return "not supported" error (not parameter validation error)
        Assert.False(result.Success);
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }


    // ========================================
    // MOVE-TO-WORKBOOK (cross-workbook) Tests
    // NOTE: Cross-workbook operations not currently supported due to COM RCW limitations
    // These tests verify graceful error handling
    // ========================================

    /// <inheritdoc/>
    [Fact]
    public void MoveToWorkbook_Default_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_Default_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_Default_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "MoveMe");

        // Act - Attempt cross-workbook move
        var result = _sheetCommands.MoveToWorkbook(sourceBatch, "MoveMe", targetBatch);

        // Assert - Should return clear error about limitation
        Assert.False(result.Success);
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("COM interop", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void MoveToWorkbook_WithBeforeSheet_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_WithBeforeSheet_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_WithBeforeSheet_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "MoveMe");
        _sheetCommands.Create(targetBatch, "Target2");

        // Act - Attempt cross-workbook move
        var result = _sheetCommands.MoveToWorkbook(sourceBatch, "MoveMe", targetBatch, beforeSheet: "Sheet1");

        // Assert - Should return clear error
        Assert.False(result.Success);
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void MoveToWorkbook_WithAfterSheet_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_WithAfterSheet_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_WithAfterSheet_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "MoveMe");
        _sheetCommands.Create(targetBatch, "Target2");

        // Act - Attempt cross-workbook move
        var result = _sheetCommands.MoveToWorkbook(sourceBatch, "MoveMe", targetBatch, afterSheet: "Target2");

        // Assert - Should return clear error
        Assert.False(result.Success);
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    [Fact]
    public void MoveToWorkbook_BothBeforeAndAfter_ReturnsNotSupportedError()
    {
        // Arrange
        var sourceFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_BothBeforeAndAfter_ReturnsNotSupportedError)}_Source", _tempDir);
        var targetFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(SheetCommandsTests), $"{nameof(MoveToWorkbook_BothBeforeAndAfter_ReturnsNotSupportedError)}_Target", _tempDir);

        using var sourceBatch = ExcelSession.BeginBatch(sourceFile);
        using var targetBatch = ExcelSession.BeginBatch(targetFile);

        _sheetCommands.Create(sourceBatch, "MoveMe");

        // Act - Attempt cross-workbook move
        var result = _sheetCommands.MoveToWorkbook(sourceBatch, "MoveMe", targetBatch, beforeSheet: "Sheet1", afterSheet: "Sheet1");

        // Assert - Should return "not supported" error (not parameter validation error)
        Assert.False(result.Success);
        Assert.Contains("not currently supported", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
}
