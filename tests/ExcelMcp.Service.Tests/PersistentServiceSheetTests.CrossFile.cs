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

        Assert.Equal(["Sheet1", "CopiedSheet"], ReadSheetNames(targetFile));
        Assert.Equal(["SourceSheet", "Sheet1"], ReadSheetNames(sourceFile));
        AssertTransferredContent(targetFile, "CopiedSheet");
        AssertTransferredContent(sourceFile, "SourceSheet");
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
    }

    [Fact]
    public void CopyToFile_NoTargetName_CopiesWithOriginalName()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "Source", "SourceSheet");
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "Target");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile));

        Assert.Equal(["Sheet1", "SourceSheet"], ReadSheetNames(targetFile));
        Assert.Equal(["SourceSheet", "Sheet1"], ReadSheetNames(sourceFile));
        AssertTransferredContent(targetFile, "SourceSheet");
        AssertTransferredContent(sourceFile, "SourceSheet");
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
    }

    [Fact]
    public void CopyToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "Source", "SourceSheet");
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "Target");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1"));

        var sheets = ReadSheetNames(targetFile);
        Assert.Equal(["Copied", "Sheet1"], sheets);
        AssertTransferredContent(targetFile, "Copied");
        Assert.Equal(["SourceSheet", "Sheet1"], ReadSheetNames(sourceFile));
        AssertTransferredContent(sourceFile, "SourceSheet");
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
    }

    [Fact]
    public void CopyToFile_WithAfterSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "Source", "SourceSheet");
        var targetFile = CreateMarkedWorkbook(nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "Target");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", afterSheet: "Sheet1"));

        var sheets = ReadSheetNames(targetFile);
        Assert.Equal(["Sheet1", "Copied"], sheets);
        AssertTransferredContent(targetFile, "Copied");
        Assert.Equal(["SourceSheet", "Sheet1"], ReadSheetNames(sourceFile));
        AssertTransferredContent(sourceFile, "SourceSheet");
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
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
        Assert.Equal(["Sheet1"], ReadSheetNames(targetFile));
        Assert.Equal(["SourceSheet", "Sheet1"], ReadSheetNames(sourceFile));
        AssertTransferredContent(sourceFile, "SourceSheet");
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
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
        Assert.Equal(["Sheet1"], ReadSheetNames(testFile));
        AssertWorkbookMarker(testFile);
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
        Assert.Equal(["Sheet1"], ReadSheetNames(targetFile));
        AssertWorkbookMarker(targetFile);
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
        Assert.Equal(["Sheet1"], ReadSheetNames(sourceFile));
        AssertWorkbookMarker(sourceFile);
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
        Assert.Equal(["SourceSheet", "Sheet1"], ReadSheetNames(sourceFile));
        Assert.Equal(["Sheet1"], ReadSheetNames(targetFile));
        AssertTransferredContent(sourceFile, "SourceSheet");
        AssertWorkbookMarker(sourceFile);
        AssertWorkbookMarker(targetFile);
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
            "Target");
        AddNamedSheetWithContent(targetFile, "CopiedSheet");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        var error = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "CopiedSheet"));

        Assert.Contains("already exists", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(["SourceSheet", "Sheet1"], ReadSheetNames(sourceFile));
        Assert.Equal(["CopiedSheet", "Sheet1"], ReadSheetNames(targetFile));
        AssertTransferredContent(sourceFile, "SourceSheet");
        AssertWorkbookMarker(targetFile);
        AssertSheetCells(targetFile, "CopiedSheet", "duplicate guard", 12, 15, "=B1+3");
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void MoveToFile_Default_MovesSheetSuccessfully()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_Default_MovesSheetSuccessfully), "Source", "MoveMe");
        var targetFile = CreateMarkedWorkbook(nameof(MoveToFile_Default_MovesSheetSuccessfully), "Target");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile));

        Assert.Equal(["Sheet1", "MoveMe"], ReadSheetNames(targetFile));
        Assert.Equal(["Sheet1"], ReadSheetNames(sourceFile));
        AssertTransferredContent(targetFile, "MoveMe");
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
    }

    [Fact]
    public void MoveToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "Source", "MoveMe");
        var targetFile = CreateMarkedWorkbook(nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "Target");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1"));

        var sheets = ReadSheetNames(targetFile);
        Assert.Equal(["MoveMe", "Sheet1"], sheets);
        AssertTransferredContent(targetFile, "MoveMe");
        Assert.Equal(["Sheet1"], ReadSheetNames(sourceFile));
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
    }

    [Fact]
    public void MoveToFile_WithAfterSheet_PositionsCorrectly()
    {
        var sourceFile = CreateWorkbookWithSheet(nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "Source", "MoveMe");
        var targetFile = CreateMarkedWorkbook(nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "Target");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, afterSheet: "Sheet1"));

        var sheets = ReadSheetNames(targetFile);
        Assert.Equal(["Sheet1", "MoveMe"], sheets);
        AssertTransferredContent(targetFile, "MoveMe");
        Assert.Equal(["Sheet1"], ReadSheetNames(sourceFile));
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
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
        Assert.Equal(["Sheet1"], ReadSheetNames(targetFile));
        Assert.Equal(["MoveMe", "Sheet1"], ReadSheetNames(sourceFile));
        AssertTransferredContent(sourceFile, "MoveMe");
        AssertWorkbookMarker(targetFile);
        AssertWorkbookMarker(sourceFile);
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
        Assert.Equal(["Sheet1"], ReadSheetNames(testFile));
        AssertWorkbookMarker(testFile);
        Assert.Equal(before, File.ReadAllBytes(testFile));
    }

    private string CreateWorkbookWithSheet(
        string scenario,
        string suffix,
        string sheetName)
    {
        var path = CreateMarkedWorkbook(scenario, suffix);
        using var batch = ExcelSession.BeginBatch(path);
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? values = null;
            Excel.Range? formula = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Add();
                sheet.Name = sheetName;
                values = sheet.Range["A1:B1"];
                values.Value2 = new object[,] { { "transferred marker", 3d } };
                formula = sheet.Range["C1"];
                formula.Formula = "=B1*2";
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

    private string CreateMarkedWorkbook(string scenario, string suffix)
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

    private static void AddNamedSheetWithContent(string workbookPath, string sheetName)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? values = null;
            Excel.Range? formula = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Add();
                sheet.Name = sheetName;
                values = sheet.Range["A1:B1"];
                values.Value2 = new object[,] { { "duplicate guard", 12d } };
                formula = sheet.Range["C1"];
                formula.Formula = "=B1+3";
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
    }

    private static void AssertWorkbookMarker(string path) =>
        AssertSheetCells(path, "Sheet1", "retained workbook marker", 71, 79, "=B1+8");

    private static void AssertTransferredContent(string workbookPath, string sheetName)
        => AssertSheetCells(workbookPath, sheetName, "transferred marker", 3, 6, "=B1*2");

    private static void AssertSheetCells(
        string workbookPath, string sheetName, string marker, double input, double calculated, string expectedFormula)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? values = null;
            Excel.Range? formula = null;
            try
            {
                sheets = context.Book.Worksheets;
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
                ComUtilities.Release(ref sheets);
            }
        });
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
