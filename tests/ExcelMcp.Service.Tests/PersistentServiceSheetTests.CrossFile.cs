using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceSheetTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void CrossFile_ReadOnlyWorkbook_RejectsBeforeChangingEitherFile(
        bool move, bool readOnlySource)
    {
        var sourceFile = CreateWorkbookWithSheet(
            nameof(CrossFile_ReadOnlyWorkbook_RejectsBeforeChangingEitherFile),
            "Source", "TransferSheet");
        var targetFile = CreateMarkedWorkbook(
            nameof(CrossFile_ReadOnlyWorkbook_RejectsBeforeChangingEitherFile), "Target");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);
        var readOnlyFile = readOnlySource ? sourceFile : targetFile;
        var originalAttributes = File.GetAttributes(readOnlyFile);
        var originalOpenHook = ExcelBatch.BeforeWorkbookOpenHook;
        try
        {
            // Change permissions after filesystem preflight so Excel determines read-only access.
            ExcelBatch.BeforeWorkbookOpenHook = (path, _) =>
            {
                if (string.Equals(path, readOnlyFile, StringComparison.OrdinalIgnoreCase))
                {
                    File.SetAttributes(readOnlyFile, originalAttributes | FileAttributes.ReadOnly);
                }
            };
            using (var batch = ExcelSession.BeginBatch(sourceFile, targetFile))
            {
                batch.Execute((_, _) =>
                {
                    Assert.Equal(readOnlySource, batch.GetWorkbook(sourceFile).ReadOnly);
                    Assert.Equal(!readOnlySource, batch.GetWorkbook(targetFile).ReadOnly);
                    Assert.True(batch.GetWorkbook(sourceFile).Saved);
                    Assert.True(batch.GetWorkbook(targetFile).Saved);
                });
            }
            File.SetAttributes(readOnlyFile, originalAttributes);

            var error = Assert.Throws<InvalidOperationException>(() =>
            {
                if (move)
                {
                    _sheetCommands.MoveToFile(sourceFile, "TransferSheet", targetFile);
                }
                else
                {
                    _sheetCommands.CopyToFile(sourceFile, "TransferSheet", targetFile);
                }
            });

            ExcelBatch.BeforeWorkbookOpenHook = originalOpenHook;
            File.SetAttributes(readOnlyFile, originalAttributes);
            Assert.Contains("Cannot change this workbook", error.Message, StringComparison.Ordinal);
            Assert.Contains("read-only", error.Message, StringComparison.Ordinal);
            Assert.Contains(
                "This operation has not changed the workbook.", error.Message, StringComparison.Ordinal);
            AssertWorkbookState(sourceFile, ["TransferSheet", "Sheet1"], "TransferSheet");
            AssertWorkbookState(targetFile, ["Sheet1"]);
            Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
            Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = originalOpenHook;
            File.SetAttributes(readOnlyFile, originalAttributes);
        }
    }

    [Fact]
    public void CopyToFile_WithTargetName_CopiesAndRenames()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_WithTargetName_CopiesAndRenames), "SourceSheet");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "CopiedSheet"));

        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1", "CopiedSheet"], "CopiedSheet");
    }

    [Fact]
    public void CopyToFile_NoTargetName_CopiesWithOriginalName()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_NoTargetName_CopiesWithOriginalName), "SourceSheet");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile));

        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1", "SourceSheet"], "SourceSheet");
    }

    [Fact]
    public void CopyToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_WithBeforeSheet_PositionsCorrectly), "SourceSheet");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", beforeSheet: "Sheet1"));

        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Copied", "Sheet1"], "Copied");
    }

    [Fact]
    public void CopyToFile_WithAfterSheet_PositionsCorrectly()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_WithAfterSheet_PositionsCorrectly), "SourceSheet");

        RequireSuccess(_sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "Copied", afterSheet: "Sheet1"));

        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1", "Copied"], "Copied");
    }

    [Fact]
    public void CopyToFile_BothBeforeAndAfter_ThrowsException()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_BothBeforeAndAfter_ThrowsException), "SourceSheet");
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
        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1"]);
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
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_InvalidTargetSheetName_RejectsBeforeChangingEitherFile), "SourceSheet");
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        var error = Assert.Throws<ArgumentException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, invalidName));

        Assert.Contains("Worksheet names must", error.Message, StringComparison.Ordinal);
        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["Sheet1"]);
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void CopyToFile_ExistingTargetSheetName_RejectsWithoutChangingEitherFile()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(CopyToFile_ExistingTargetSheetName_RejectsWithoutChangingEitherFile),
            "SourceSheet", "CopiedSheet", targetDuplicateGuard: true);
        var sourceBytes = File.ReadAllBytes(sourceFile);
        var targetBytes = File.ReadAllBytes(targetFile);

        var error = Assert.Throws<InvalidOperationException>(() =>
            _sheetCommands.CopyToFile(sourceFile, "SourceSheet", targetFile, "CopiedSheet"));

        Assert.Contains("already exists", error.Message, StringComparison.OrdinalIgnoreCase);
        AssertWorkbookStates(
            sourceFile, ["SourceSheet", "Sheet1"], "SourceSheet",
            targetFile, ["CopiedSheet", "Sheet1"], "CopiedSheet", targetDuplicateGuard: true);
        Assert.Equal(sourceBytes, File.ReadAllBytes(sourceFile));
        Assert.Equal(targetBytes, File.ReadAllBytes(targetFile));
    }

    [Fact]
    public void MoveToFile_Default_MovesSheetSuccessfully()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(MoveToFile_Default_MovesSheetSuccessfully), "MoveMe");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile));

        AssertWorkbookStates(
            sourceFile, ["Sheet1"], null,
            targetFile, ["Sheet1", "MoveMe"], "MoveMe");
    }

    [Fact]
    public void MoveToFile_WithBeforeSheet_PositionsCorrectly()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(MoveToFile_WithBeforeSheet_PositionsCorrectly), "MoveMe");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, beforeSheet: "Sheet1"));

        AssertWorkbookStates(
            sourceFile, ["Sheet1"], null,
            targetFile, ["MoveMe", "Sheet1"], "MoveMe");
    }

    [Fact]
    public void MoveToFile_WithAfterSheet_PositionsCorrectly()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(MoveToFile_WithAfterSheet_PositionsCorrectly), "MoveMe");

        RequireSuccess(_sheetCommands.MoveToFile(sourceFile, "MoveMe", targetFile, afterSheet: "Sheet1"));

        AssertWorkbookStates(
            sourceFile, ["Sheet1"], null,
            targetFile, ["Sheet1", "MoveMe"], "MoveMe");
    }

    [Fact]
    public void MoveToFile_BothBeforeAndAfter_ThrowsException()
    {
        var (sourceFile, targetFile) = CreateMarkedWorkbookPair(
            nameof(MoveToFile_BothBeforeAndAfter_ThrowsException), "MoveMe");
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
        AssertWorkbookStates(
            sourceFile, ["MoveMe", "Sheet1"], "MoveMe",
            targetFile, ["Sheet1"]);
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

    private (string Source, string Target) CreateMarkedWorkbookPair(
        string scenario,
        string sourceSheet,
        string? targetSheet = null,
        bool targetDuplicateGuard = false)
    {
        var sourceFile = _fixture.CreateBlankWorkbook(scenario, "Source");
        var targetFile = _fixture.CreateBlankWorkbook(scenario, "Target");
        using var batch = ExcelSession.BeginBatch(sourceFile, targetFile);
        batch.Execute((_, _) =>
        {
            SeedWorkbook(batch.GetWorkbook(sourceFile), sourceSheet);
            SeedWorkbook(batch.GetWorkbook(targetFile), targetSheet, targetDuplicateGuard);
            batch.GetWorkbook(targetFile).Save();
        });
        batch.Save();
        return (sourceFile, targetFile);
    }

    private string CreateMarkedWorkbook(
        string scenario, string suffix, string? additionalSheet = null, bool duplicateGuard = false)
    {
        var path = _fixture.CreateBlankWorkbook(scenario, suffix);
        using var batch = ExcelSession.BeginBatch(path);
        batch.Execute((context, _) => SeedWorkbook(context.Book, additionalSheet, duplicateGuard));
        batch.Save();
        return path;
    }

    private static void SeedWorkbook(
        Excel.Workbook workbook,
        string? additionalSheet,
        bool duplicateGuard = false)
    {
        Excel.Sheets? sheets = null;
        Excel.Worksheet? sheet = null;
        Excel.Range? values = null;
        Excel.Range? formula = null;
        try
        {
            sheets = workbook.Worksheets;
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
    }

    private static void AssertWorkbookState(
        string workbookPath, string[] expectedNames, string? additionalSheet = null, bool duplicateGuard = false)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) => AssertWorkbookState(
            context.Book, expectedNames, additionalSheet, duplicateGuard));
    }

    private static void AssertWorkbookStates(
        string sourcePath,
        string[] sourceExpectedNames,
        string? sourceSheet,
        string targetPath,
        string[] targetExpectedNames,
        string? targetSheet = null,
        bool targetDuplicateGuard = false)
    {
        using var batch = ExcelSession.BeginBatch(sourcePath, targetPath);
        batch.Execute((_, _) =>
        {
            AssertWorkbookState(batch.GetWorkbook(sourcePath), sourceExpectedNames, sourceSheet);
            AssertWorkbookState(batch.GetWorkbook(targetPath), targetExpectedNames, targetSheet, targetDuplicateGuard);
        });
    }

    private static void AssertWorkbookState(
        Excel.Workbook workbook,
        string[] expectedNames,
        string? additionalSheet = null,
        bool duplicateGuard = false)
    {
        Excel.Sheets? sheets = null;
        try
        {
            sheets = workbook.Worksheets;
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
