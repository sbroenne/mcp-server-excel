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
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_WithTargetName_CopiesAndRenames), "Target");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "CopiedSheet"));

        AssertWorkbookState(targetFile, ["Sheet1", "CopiedSheet"], "CopiedSheet");
        AssertWorkbookState(sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet");
    }

    [Fact]
    public void CopyToFile_NoTargetName_CopiesWithOriginalName()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "Source", "SourceSheet");
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "Target");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile));

        AssertWorkbookState(targetFile, ["Sheet1", "SourceSheet"], "SourceSheet");
        AssertWorkbookState(sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet");
    }

    [Fact]
    public void CopyToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "Source", "SourceSheet");
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "Target");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1"));

        AssertWorkbookState(targetFile, ["Copied", "Sheet1"], "Copied");
        AssertWorkbookState(sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet");
    }

    [Fact]
    public void CopyToFile_WithAfterSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "Source", "SourceSheet");
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "Target");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", afterSheet: "Sheet1"));

        AssertWorkbookState(targetFile, ["Sheet1", "Copied"], "Copied");
        AssertWorkbookState(sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet");
    }

    [Fact]
    public void CopyToFile_BothBeforeAndAfter_ThrowsException()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_BothBeforeAndAfter_ThrowsException), "Source", "SourceSheet");
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_BothBeforeAndAfter_ThrowsException), "Target");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

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
        AssertWorkbookState(targetFile, ["Sheet1"]);
        AssertWorkbookState(sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet");
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void CopyToFile_SameFile_ThrowsException()
    {
        var testFile = CreateMarkedWorkbook(nameof(CopyToFile_SameFile_ThrowsException), "Test");
        var before = File.ReadAllBytes(testFile);

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.CopyToFile(testFile, "Sheet1", testFile, "Copied"));
        Assert.Contains("same-file copy", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertWorkbookState(testFile, ["Sheet1"]);
        Assert.Equal(before, File.ReadAllBytes(testFile));
    }

    [Fact]
    public void CopyToFile_SourceFileNotFound_ThrowsException()
    {
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_SourceFileNotFound_ThrowsException), "Target");
        var before = File.ReadAllBytes(targetFile);
        var nonExistentSource = Path.Combine(Path.GetDirectoryName(targetFile)!, "NonExistent.xlsx");

        var exception = Assert.Throws<FileNotFoundException>(() =>
            _sheetCommands.CopyToFile(nonExistentSource, "Sheet1", targetFile));
        Assert.Contains("Source file not found", exception.Message);
        AssertWorkbookState(targetFile, ["Sheet1"]);
        Assert.Equal(before, File.ReadAllBytes(targetFile));
        Assert.False(File.Exists(nonExistentSource));
    }

    [Fact]
    public void CopyToFile_TargetFileNotFound_ThrowsException()
    {
        var sourceFile = CreateMarkedWorkbook(nameof(CopyToFile_TargetFileNotFound_ThrowsException), "Source");
        var before = File.ReadAllBytes(sourceFile);
        var nonExistentTarget = Path.Combine(Path.GetDirectoryName(sourceFile)!, "NonExistent.xlsx");

        var exception = Assert.Throws<FileNotFoundException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "Sheet1", nonExistentTarget));
        Assert.Contains("Target file not found", exception.Message);
        AssertWorkbookState(sourceFile, ["Sheet1"]);
        Assert.Equal(before, File.ReadAllBytes(sourceFile));
        Assert.False(File.Exists(nonExistentTarget));
    }

    [Theory]
    [InlineData("History")]
    [InlineData("Bad/Name")]
    [InlineData("12345678901234567890123456789012")]
    [InlineData("   ")]
    public void CopyToFile_InvalidTargetSheetName_RejectsBeforeChangingEitherFile(string invalidName)
    {
        var sourceFile = CreateWorkbookWithSheet(
            nameof(CopyToFile_InvalidTargetSheetName_RejectsBeforeChangingEitherFile),
            "Source", "SourceSheet");
        var targetFile = CreateMarkedWorkbook(
            nameof(CopyToFile_InvalidTargetSheetName_RejectsBeforeChangingEitherFile),
            "Target");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        var error = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, invalidName));

        Assert.Contains("Worksheet names must", error.Message, StringComparison.Ordinal);
        AssertWorkbookState(sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet");
        AssertWorkbookState(targetFile, ["Sheet1"]);
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void CopyToFile_ExistingTargetSheetName_RejectsWithoutChangingEitherFile()
    {
        var sourceFile = CreateWorkbookWithSheet(
            nameof(CopyToFile_ExistingTargetSheetName_RejectsWithoutChangingEitherFile),
            "Source",
            "SourceSheet");
        var targetFile = CreateMarkedWorkbook(
            nameof(CopyToFile_ExistingTargetSheetName_RejectsWithoutChangingEitherFile),
            "Target", "CopiedSheet", duplicateGuard: true);
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        var error = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "CopiedSheet"));

        Assert.Contains("already exists", error.Message, StringComparison.OrdinalIgnoreCase);
        AssertWorkbookState(sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet");
        AssertWorkbookState(targetFile, ["CopiedSheet", "Sheet1"], "CopiedSheet", duplicateGuard: true);
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void MoveToFile_Default_MovesSheetSuccessfully()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_Default_MovesSheetSuccessfully), "Source", "MoveMe");
        var targetFile = CreateMarkedWorkbook(nameof(MoveToFile_Default_MovesSheetSuccessfully), "Target");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile));

        AssertWorkbookState(targetFile, ["Sheet1", "MoveMe"], "MoveMe");
        AssertWorkbookState(sourceFile, ["Sheet1"]);
    }

    [Fact]
    public void MoveToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "Source", "MoveMe");
        var targetFile = CreateMarkedWorkbook(nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "Target");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1"));

        AssertWorkbookState(targetFile, ["MoveMe", "Sheet1"], "MoveMe");
        AssertWorkbookState(sourceFile, ["Sheet1"]);
    }

    [Fact]
    public void MoveToFile_WithAfterSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "Source", "MoveMe");
        var targetFile = CreateMarkedWorkbook(nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "Target");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, afterSheet: "Sheet1"));

        AssertWorkbookState(targetFile, ["Sheet1", "MoveMe"], "MoveMe");
        AssertWorkbookState(sourceFile, ["Sheet1"]);
    }

    [Fact]
    public void MoveToFile_BothBeforeAndAfter_ThrowsException()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_BothBeforeAndAfter_ThrowsException), "Source", "MoveMe");
        var targetFile = CreateMarkedWorkbook(nameof(MoveToFile_BothBeforeAndAfter_ThrowsException), "Target");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

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
        AssertWorkbookState(targetFile, ["Sheet1"]);
        AssertWorkbookState(sourceFile, ["MoveMe", "Sheet1"], "MoveMe");
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void MoveToFile_SameFile_ThrowsException()
    {
        var testFile = CreateMarkedWorkbook(nameof(MoveToFile_SameFile_ThrowsException), "Test");
        var before = File.ReadAllBytes(testFile);

        var exception = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.MoveToFile(testFile, "Sheet1", testFile));
        Assert.Contains("same-file move", exception.Message, StringComparison.OrdinalIgnoreCase);
        AssertWorkbookState(testFile, ["Sheet1"]);
        Assert.Equal(before, File.ReadAllBytes(testFile));
    }

    private string CreateWorkbookWithSheet(
        string scenario,
        string suffix,
        string sheetName) => CreateMarkedWorkbook(scenario, suffix, sheetName);

    private string CreateMarkedWorkbook(
        string scenario, string suffix, string? additionalSheet = null, bool duplicateGuard = false)
    {
        var path = _fixture.CreateBlankWorkbook(scenario, suffix);
        using var batch = ExcelSession.BeginBatch(path);
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? values = null;
            Excel.Range? formula = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item["Sheet1"];
                values = sheet.Range["A1:B1"];
                values.Value2 = new object[,] { { "retained workbook marker", 71d } };
                formula = sheet.Range["C1"];
                formula.Formula = "=B1+8";
                if (additionalSheet is not null)
                {
                    ComUtilities.Release(ref formula);
                    ComUtilities.Release(ref values);
                    ComUtilities.Release(ref sheet);
                    sheet = (Excel.Worksheet)sheets.Add();
                    sheet.Name = additionalSheet;
                    values = sheet.Range["A1:B1"];
                    values.Value2 = duplicateGuard
                        ? new object[,] { { "duplicate guard", 12d } }
                        : new object[,] { { "transferred marker", 3d } };
                    formula = sheet.Range["C1"];
                    formula.Formula = duplicateGuard ? "=B1+3" : "=B1*2";
                }
            }
            finally
            {
                ComUtilities.Release(ref formula);
                ComUtilities.Release(ref values);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        batch.Save();
        return path;
    }

    private static void AssertWorkbookState(
        string workbookPath, string[] expectedNames, string? additionalSheet = null, bool duplicateGuard = false)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            try
            {
                sheets = context.Book.Worksheets;
                var names = new List<string>();
                for (var index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet? sheet = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[index];
                        names.Add(sheet.Name);
                    }
                    finally { ComUtilities.Release(ref sheet); }
                }
                Assert.Equal(expectedNames, names);
                AssertSheetCells(sheets, "Sheet1", "retained workbook marker", 71, 79, "=B1+8");
                if (additionalSheet is not null)
                {
                    if (duplicateGuard)
                    {
                        AssertSheetCells(sheets, additionalSheet, "duplicate guard", 12, 15, "=B1+3");
                    }
                    else
                    {
                        AssertSheetCells(sheets, additionalSheet, "transferred marker", 3, 6, "=B1*2");
                    }
                }
            }
            finally { ComUtilities.Release(ref sheets); }
        });
    }

    private static void AssertSheetCells(
        Excel.Sheets sheets, string sheetName, string marker, double input, double calculated, string expectedFormula)
    {
        Excel.Worksheet? sheet = null;
        Excel.Range? values = null;
        Excel.Range? formula = null;
        try
        {
            sheet = (Excel.Worksheet)sheets[sheetName];
            values = sheet.Range["A1:C1"];
            var cells = Assert.IsType<object[,]>(values.Value2);
            Assert.Equal(1, cells.GetLength(0));
            Assert.Equal(3, cells.GetLength(1));
            Assert.Equal(marker, cells[1, 1]);
            Assert.Equal(input, Convert.ToDouble(cells[1, 2], System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(calculated, Convert.ToDouble(cells[1, 3], System.Globalization.CultureInfo.InvariantCulture));
            formula = sheet.Range["C1"];
            Assert.Equal(expectedFormula, formula.Formula);
        }
        finally
        {
            ComUtilities.Release(ref formula);
            ComUtilities.Release(ref values);
            ComUtilities.Release(ref sheet);
        }
    }
}
