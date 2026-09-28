using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Fact]
    public void CopyToFile_WithTargetName_CopiesAndRenames()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_WithTargetName_CopiesAndRenames), "Source", "SourceSheet");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(CopyToFile_WithTargetName_CopiesAndRenames), "Target");

        _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "CopiedSheet");

        Assert.Contains("CopiedSheet", ReadSheetNames(targetFile));
        Assert.Contains("SourceSheet", ReadSheetNames(sourceFile));
    }

    [Fact]
    public void CopyToFile_NoTargetName_CopiesWithOriginalName()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "Source", "SourceSheet");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "Target");

        _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile);

        Assert.Contains("SourceSheet", ReadSheetNames(targetFile));
    }

    [Fact]
    public void CopyToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "Source", "SourceSheet");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "Target");

        _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1");

        var sheets = ReadSheetNames(targetFile);
        var copiedIndex = sheets.IndexOf("Copied");
        var sheet1Index = sheets.IndexOf("Sheet1");
        Assert.True(
            copiedIndex < sheet1Index,
            $"Expected Copied (index {copiedIndex}) to be before Sheet1 (index {sheet1Index})");
    }

    [Fact]
    public void CopyToFile_WithAfterSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "Source", "SourceSheet");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "Target");

        _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", afterSheet: "Sheet1");

        var sheets = ReadSheetNames(targetFile);
        var copiedIndex = sheets.IndexOf("Copied");
        var sheet1Index = sheets.IndexOf("Sheet1");
        Assert.True(
            copiedIndex > sheet1Index,
            $"Expected Copied (index {copiedIndex}) to be after Sheet1 (index {sheet1Index})");
    }

    [Fact]
    public void CopyToFile_BothBeforeAndAfter_ThrowsException()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_BothBeforeAndAfter_ThrowsException), "Source", "SourceSheet");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(CopyToFile_BothBeforeAndAfter_ThrowsException), "Target");

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.CopyToFile(
                sourceFile,
                "SourceSheet",
                targetFile,
                "Copied",
                beforeSheet: "Sheet1",
                afterSheet: "Sheet1"));
        Assert.Contains(
            "both beforeSheet and afterSheet",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CopyToFile_SameFile_ThrowsException()
    {
        var testFile = _fixture.CreateBlankWorkbook(nameof(CopyToFile_SameFile_ThrowsException), "Test");

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.CopyToFile(testFile, "Sheet1", testFile, "Copied"));
        Assert.Contains("same-file copy", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void CopyToFile_SourceFileNotFound_ThrowsException()
    {
        var targetFile = _fixture.CreateBlankWorkbook(nameof(CopyToFile_SourceFileNotFound_ThrowsException), "Target");
        var nonExistentSource = Path.Combine(Path.GetDirectoryName(targetFile)!, "NonExistent.xlsx");

        var exception = Assert.Throws<FileNotFoundException>(() =>
            _sheetCommands.CopyToFile(nonExistentSource, "Sheet1", targetFile));
        Assert.Contains("Source file not found", exception.Message);
    }

    [Fact]
    public void CopyToFile_TargetFileNotFound_ThrowsException()
    {
        var sourceFile = _fixture.CreateBlankWorkbook(nameof(CopyToFile_TargetFileNotFound_ThrowsException), "Source");
        var nonExistentTarget = Path.Combine(Path.GetDirectoryName(sourceFile)!, "NonExistent.xlsx");

        var exception = Assert.Throws<FileNotFoundException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "Sheet1", nonExistentTarget));
        Assert.Contains("Target file not found", exception.Message);
    }

    [Fact]
    public void MoveToFile_Default_MovesSheetSuccessfully()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_Default_MovesSheetSuccessfully), "Source", "MoveMe");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(MoveToFile_Default_MovesSheetSuccessfully), "Target");

        _sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile);

        Assert.Contains("MoveMe", ReadSheetNames(targetFile));
        Assert.DoesNotContain("MoveMe", ReadSheetNames(sourceFile));
    }

    [Fact]
    public void MoveToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "Source", "MoveMe");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "Target");

        _sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1");

        var sheets = ReadSheetNames(targetFile);
        var moveMeIndex = sheets.IndexOf("MoveMe");
        var sheet1Index = sheets.IndexOf("Sheet1");
        Assert.True(
            moveMeIndex < sheet1Index,
            $"Expected MoveMe (index {moveMeIndex}) to be before Sheet1 (index {sheet1Index})");
    }

    [Fact]
    public void MoveToFile_WithAfterSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "Source", "MoveMe");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "Target");

        _sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, afterSheet: "Sheet1");

        var sheets = ReadSheetNames(targetFile);
        var moveMeIndex = sheets.IndexOf("MoveMe");
        var sheet1Index = sheets.IndexOf("Sheet1");
        Assert.True(
            moveMeIndex > sheet1Index,
            $"Expected MoveMe (index {moveMeIndex}) to be after Sheet1 (index {sheet1Index})");
    }

    [Fact]
    public void MoveToFile_BothBeforeAndAfter_ThrowsException()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_BothBeforeAndAfter_ThrowsException), "Source", "MoveMe");
        var targetFile = _fixture.CreateBlankWorkbook(nameof(MoveToFile_BothBeforeAndAfter_ThrowsException), "Target");

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.MoveToFile(
                sourceFile,
                "MoveMe",
                targetFile,
                beforeSheet: "Sheet1",
                afterSheet: "Sheet1"));
        Assert.Contains(
            "both beforeSheet and afterSheet",
            exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MoveToFile_SameFile_ThrowsException()
    {
        var testFile = _fixture.CreateBlankWorkbook(nameof(MoveToFile_SameFile_ThrowsException), "Test");

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.MoveToFile(testFile, "Sheet1", testFile));
        Assert.Contains("same-file move", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private string CreateWorkbookWithSheet(
        string scenario,
        string suffix,
        string sheetName)
    {
        var path = _fixture.CreateBlankWorkbook(scenario, suffix);
        using var batch = ExcelSession.BeginBatch(path);
        batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets.Add();
                sheet.Name = sheetName;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
        batch.Save();
        return path;
    }

    private static List<string> ReadSheetNames(string workbookPath)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        return batch.Execute((ctx, ct) =>
        {
            var names = new List<string>();
            Excel.Sheets? sheets = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                for (var index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet? sheet = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[index];
                        names.Add(sheet.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref sheet);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }

            return names;
        });
    }
}
