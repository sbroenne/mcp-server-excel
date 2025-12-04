using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Tests for Sheet move and cross-file operations (move, copy-to-file, move-to-file)
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
        // Arrange - Move tests need isolated file to control sheet order
        var testFile = _fixture.CreateCrossWorkbookTestFile(nameof(Move_WithBeforeSheet_RepositionsSheet));

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "MoveMe");
        _sheetCommands.Create(batch, "Target");

        // Get initial position of MoveMe
        _sheetCommands.List(batch);

        // Act - Move MoveMe before Sheet1
        _sheetCommands.Move(batch, "MoveMe", beforeSheet: "Sheet1");  // Move throws on error

        // Assert - reaching here means move succeeded

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
        // Arrange - Move tests need isolated file to control sheet order
        var testFile = _fixture.CreateCrossWorkbookTestFile(nameof(Move_WithAfterSheet_RepositionsSheet));

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "MoveMe");
        _sheetCommands.Create(batch, "Target");

        // Get initial position
        _sheetCommands.List(batch);

        // Act - Move MoveMe after Target
        _sheetCommands.Move(batch, "MoveMe", afterSheet: "Target");  // Move throws on error

        // Assert - reaching here means move succeeded

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
        // Arrange - Move tests need isolated file to control sheet order
        var testFile = _fixture.CreateCrossWorkbookTestFile(nameof(Move_NoPositionSpecified_MovesToEnd));

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Sheet2");
        _sheetCommands.Create(batch, "Sheet3");

        // Act - Move Sheet1 without specifying position
        _sheetCommands.Move(batch, "Sheet1");  // Move throws on error

        // Assert - reaching here means move succeeded

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
        var testFile = _fixture.CreateCrossWorkbookTestFile(nameof(Move_BothBeforeAndAfter_ThrowsException));

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
        var testFile = _fixture.CreateCrossWorkbookTestFile(nameof(Move_NonExistentSheet_ThrowsException));

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
        var testFile = _fixture.CreateCrossWorkbookTestFile(nameof(Move_NonExistentTargetSheet_ThrowsException));

        using var batch = ExcelSession.BeginBatch(testFile);
        _sheetCommands.Create(batch, "Sheet2");

        // Act & Assert - Should throw when target sheet doesn't exist
        var exception = Assert.Throws<InvalidOperationException>(
            () => _sheetCommands.Move(batch, "Sheet2", beforeSheet: "NonExistent"));
        Assert.Contains("not found", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ========================================
    // COPY-TO-FILE (atomic cross-file) Tests
    // Tests for copying sheets between different files using atomic operations
    // ========================================

    [Fact]
    public void CopyToFile_WithTargetName_CopiesAndRenames()
    {
        // Arrange - Create source and target files with a sheet to copy
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_WithTargetName_CopiesAndRenames), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_WithTargetName_CopiesAndRenames), "Target");

        // Create source sheet using a batch
        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "SourceSheet");
            batch.Save();
        }

        // Act - Copy sheet to target file with new name (atomic operation)
        _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "CopiedSheet");

        // Assert - Verify sheet exists in target file with new name
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            var targetList = _sheetCommands.List(batch);
            Assert.Contains(targetList.Worksheets, s => s.Name == "CopiedSheet");
        }

        // Verify source sheet still exists in source file
        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            var sourceList = _sheetCommands.List(batch);
            Assert.Contains(sourceList.Worksheets, s => s.Name == "SourceSheet");
        }
    }

    [Fact]
    public void CopyToFile_NoTargetName_CopiesWithOriginalName()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "Target");

        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "SourceSheet");
            batch.Save();
        }

        // Act - Copy without specifying target name
        _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile);

        // Assert - Verify sheet was copied with original name
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            var targetList = _sheetCommands.List(batch);
            Assert.Contains(targetList.Worksheets, s => s.Name == "SourceSheet");
        }
    }

    [Fact]
    public void CopyToFile_WithBeforeSheet_PositionsCorrectly()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "Target");

        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "SourceSheet");
            batch.Save();
        }

        // Act - Copy before Sheet1 in target file
        _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1");

        // Assert - Verify sheet was copied and positioned before Sheet1
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            var targetList = _sheetCommands.List(batch);
            var sheets = targetList.Worksheets.ToList();
            var copiedIndex = sheets.FindIndex(s => s.Name == "Copied");
            var sheet1Index = sheets.FindIndex(s => s.Name == "Sheet1");
            Assert.True(copiedIndex < sheet1Index, $"Expected Copied (index {copiedIndex}) to be before Sheet1 (index {sheet1Index})");
        }
    }

    [Fact]
    public void CopyToFile_WithAfterSheet_PositionsCorrectly()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "Target");

        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "SourceSheet");
            batch.Save();
        }

        // Act - Copy after Sheet1 in target file
        _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", afterSheet: "Sheet1");

        // Assert - Verify sheet was copied and positioned after Sheet1
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            var targetList = _sheetCommands.List(batch);
            var sheets = targetList.Worksheets.ToList();
            var copiedIndex = sheets.FindIndex(s => s.Name == "Copied");
            var sheet1Index = sheets.FindIndex(s => s.Name == "Sheet1");
            Assert.True(copiedIndex > sheet1Index, $"Expected Copied (index {copiedIndex}) to be after Sheet1 (index {sheet1Index})");
        }
    }

    [Fact]
    public void CopyToFile_BothBeforeAndAfter_ThrowsException()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_BothBeforeAndAfter_ThrowsException), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_BothBeforeAndAfter_ThrowsException), "Target");

        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "SourceSheet");
            batch.Save();
        }

        // Act & Assert - Should throw ArgumentException when both beforeSheet and afterSheet specified
        var exception = Assert.Throws<ArgumentException>(
            () => _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1", afterSheet: "Sheet1"));
        Assert.Contains("both beforeSheet and afterSheet", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CopyToFile_SameFile_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_SameFile_ThrowsException), "Test");

        // Act & Assert - Should throw when source and target are the same
        var exception = Assert.Throws<ArgumentException>(
            () => _sheetCommands.CopyToFile(testFile, "Sheet1", testFile, "Copied"));
        Assert.Contains("same-file copy", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CopyToFile_SourceFileNotFound_ThrowsException()
    {
        // Arrange
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_SourceFileNotFound_ThrowsException), "Target");
        var nonExistentSource = Path.Combine(_fixture.TempDir, "NonExistent.xlsx");

        // Act & Assert
        var exception = Assert.Throws<FileNotFoundException>(
            () => _sheetCommands.CopyToFile(nonExistentSource, "Sheet1", targetFile));
        Assert.Contains("Source file not found", exception.Message);
    }

    [Fact]
    public void CopyToFile_TargetFileNotFound_ThrowsException()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(CopyToFile_TargetFileNotFound_ThrowsException), "Source");
        var nonExistentTarget = Path.Combine(_fixture.TempDir, "NonExistent.xlsx");

        // Act & Assert
        var exception = Assert.Throws<FileNotFoundException>(
            () => _sheetCommands.CopyToFile(sourceFile, "Sheet1", nonExistentTarget));
        Assert.Contains("Target file not found", exception.Message);
    }

    // ========================================
    // MOVE-TO-FILE (atomic cross-file) Tests
    // Tests for moving sheets between different files using atomic operations
    // ========================================

    [Fact]
    public void MoveToFile_Default_MovesSheetSuccessfully()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_Default_MovesSheetSuccessfully), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_Default_MovesSheetSuccessfully), "Target");

        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "MoveMe");
            batch.Save();
        }

        // Act - Move sheet to target file (atomic operation)
        _sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile);

        // Assert - Verify sheet exists in target file
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            var targetList = _sheetCommands.List(batch);
            Assert.Contains(targetList.Worksheets, s => s.Name == "MoveMe");
        }

        // Verify sheet no longer exists in source file
        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            var sourceList = _sheetCommands.List(batch);
            Assert.DoesNotContain(sourceList.Worksheets, s => s.Name == "MoveMe");
        }
    }

    [Fact]
    public void MoveToFile_WithBeforeSheet_PositionsCorrectly()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "Target");

        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "MoveMe");
            batch.Save();
        }

        // Act - Move before Sheet1 in target file
        _sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1");

        // Assert - Verify sheet was moved and positioned correctly
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            var targetList = _sheetCommands.List(batch);
            var sheets = targetList.Worksheets.ToList();
            var moveMeIndex = sheets.FindIndex(s => s.Name == "MoveMe");
            var sheet1Index = sheets.FindIndex(s => s.Name == "Sheet1");
            Assert.True(moveMeIndex < sheet1Index, $"Expected MoveMe (index {moveMeIndex}) to be before Sheet1 (index {sheet1Index})");
        }
    }

    [Fact]
    public void MoveToFile_WithAfterSheet_PositionsCorrectly()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "Target");

        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "MoveMe");
            batch.Save();
        }

        // Act - Move after Sheet1 in target file
        _sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, afterSheet: "Sheet1");

        // Assert - Verify sheet was moved and positioned correctly
        using (var batch = ExcelSession.BeginBatch(targetFile))
        {
            var targetList = _sheetCommands.List(batch);
            var sheets = targetList.Worksheets.ToList();
            var moveMeIndex = sheets.FindIndex(s => s.Name == "MoveMe");
            var sheet1Index = sheets.FindIndex(s => s.Name == "Sheet1");
            Assert.True(moveMeIndex > sheet1Index, $"Expected MoveMe (index {moveMeIndex}) to be after Sheet1 (index {sheet1Index})");
        }
    }

    [Fact]
    public void MoveToFile_BothBeforeAndAfter_ThrowsException()
    {
        // Arrange
        var sourceFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_BothBeforeAndAfter_ThrowsException), "Source");
        var targetFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_BothBeforeAndAfter_ThrowsException), "Target");

        using (var batch = ExcelSession.BeginBatch(sourceFile))
        {
            _sheetCommands.Create(batch, "MoveMe");
            batch.Save();
        }

        // Act & Assert
        var exception = Assert.Throws<ArgumentException>(
            () => _sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1", afterSheet: "Sheet1"));
        Assert.Contains("both beforeSheet and afterSheet", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MoveToFile_SameFile_ThrowsException()
    {
        // Arrange
        var testFile = _fixture.CreateCrossWorkbookTestFile(nameof(MoveToFile_SameFile_ThrowsException), "Test");

        // Act & Assert - Should throw when source and target are the same
        var exception = Assert.Throws<ArgumentException>(
            () => _sheetCommands.MoveToFile(testFile, "Sheet1", testFile));
        Assert.Contains("same-file move", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
