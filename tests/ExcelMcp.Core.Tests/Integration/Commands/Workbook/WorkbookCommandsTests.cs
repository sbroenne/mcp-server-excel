using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Workbook;

/// <summary>
/// Integration tests for workbook-level lifecycle and metadata operations.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "true")]
public class WorkbookCommandsTests : IClassFixture<FileTestsFixture>
{
    private readonly WorkbookCommands _commands = new();
    private readonly FileTestsFixture _fixture;

    public WorkbookCommandsTests(FileTestsFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void SetProtection_ProtectsAndUnprotectsWorkbook()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var protectResult = _commands.SetProtection(batch, true);
        Assert.True(protectResult.Success, $"Expected protect to succeed but got error: {protectResult.ErrorMessage}");
        Assert.True(IsWorkbookProtected(batch));

        var unprotectResult = _commands.SetProtection(batch, false);
        Assert.True(unprotectResult.Success, $"Expected unprotect to succeed but got error: {unprotectResult.ErrorMessage}");
        Assert.False(IsWorkbookProtected(batch));
    }

    [Fact]
    public void SetViewOptions_UpdatesGridlinesAndHeadings()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var setResult = _commands.SetViewOptions(batch, displayGridlines: false, displayHeadings: true);
        Assert.True(setResult.Success, $"Expected view options to update but got error: {setResult.ErrorMessage}");

        var getResult = _commands.GetViewOptions(batch);
        Assert.True(getResult.Success, $"Expected view options to be read back but got error: {getResult.ErrorMessage}");
        Assert.False(getResult.DisplayGridlines);
        Assert.True(getResult.DisplayHeadings);
    }

    [Fact]
    public void GetInfo_ReturnsActiveWorkbookMetadata()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var result = _commands.GetInfo(batch);

        Assert.True(result.Success);
        Assert.Equal(Path.GetFileName(testFile), result.Name);
        Assert.Equal(Path.GetFullPath(testFile), result.FullName, ignoreCase: true);
        Assert.Equal("xlsx", result.Format);
        Assert.False(result.ReadOnly);
    }

    [Fact]
    public void CustomDocumentProperty_CrudRoundTrip_PreservesValue()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var setResult = _commands.SetDocumentProperty(
            batch,
            "AutomationTag",
            "alpha",
            DocumentPropertyScope.Custom);
        var getResult = _commands.GetDocumentProperty(
            batch,
            "AutomationTag",
            DocumentPropertyScope.Custom);
        var listResult = _commands.ListDocumentProperties(batch, includeBuiltIn: false, includeCustom: true);
        var deleteResult = _commands.DeleteDocumentProperty(batch, "AutomationTag");

        Assert.True(setResult.Success);
        Assert.True(getResult.Success);
        Assert.Equal("alpha", getResult.Property.Value);
        Assert.Contains(listResult.Properties, property =>
            property.Name == "AutomationTag" &&
            property.Value == "alpha" &&
            property.Scope == "custom");
        Assert.True(deleteResult.Success);
        Assert.Throws<InvalidOperationException>(() =>
            _commands.GetDocumentProperty(batch, "AutomationTag", DocumentPropertyScope.Custom));
    }

    [Fact]
    public void BuiltInDocumentProperty_SetAndGet_UpdatesTitle()
    {
        var testFile = _fixture.CreateTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        var setResult = _commands.SetDocumentProperty(
            batch,
            "Title",
            "Quarterly workbook",
            DocumentPropertyScope.BuiltIn);
        var getResult = _commands.GetDocumentProperty(
            batch,
            "Title",
            DocumentPropertyScope.BuiltIn);

        Assert.True(setResult.Success);
        Assert.Equal("Quarterly workbook", getResult.Property.Value);
        Assert.Equal("built-in", getResult.Property.Scope);
    }

    [Fact]
    public void SaveCopyAs_CreatesIndependentWorkbookCopy()
    {
        var testFile = _fixture.CreateTestFile();
        var copyPath = Path.Join(_fixture.TempDir, $"WorkbookCopy_{Guid.NewGuid():N}.xlsx");
        using var batch = ExcelSession.BeginBatch(testFile);

        var result = _commands.SaveCopyAs(batch, copyPath, overwrite: false);

        Assert.True(result.Success);
        Assert.True(System.IO.File.Exists(copyPath));
        Assert.Equal(Path.GetFullPath(testFile), batch.WorkbookPath, ignoreCase: true);
    }

    [Theory]
    [InlineData(WorkbookSaveFormat.Xlsx, "xlsx")]
    [InlineData(WorkbookSaveFormat.Xlsm, "xlsm")]
    [InlineData(WorkbookSaveFormat.Xlsb, "xlsb")]
    [InlineData(WorkbookSaveFormat.Xls, "xls")]
    public void SaveAs_ChangesWorkbookFormatAndSessionPath(
        WorkbookSaveFormat format,
        string extension)
    {
        var testFile = _fixture.CreateTestFile();
        var outputPath = Path.Join(_fixture.TempDir, $"WorkbookSaveAs_{Guid.NewGuid():N}.{extension}");
        using var batch = ExcelSession.BeginBatch(testFile);

        var result = _commands.SaveAs(batch, outputPath, format, overwrite: false);
        var info = _commands.GetInfo(batch);
        var contextPath = batch.Execute((context, _) => context.WorkbookPath);

        Assert.True(result.Success);
        Assert.True(System.IO.File.Exists(outputPath));
        Assert.Equal(Path.GetFullPath(outputPath), batch.WorkbookPath, ignoreCase: true);
        Assert.Equal(Path.GetFullPath(outputPath), contextPath, ignoreCase: true);
        Assert.Equal(Path.GetFullPath(outputPath), info.FullName, ignoreCase: true);
        Assert.Equal(extension, info.Format);
    }

    [Theory]
    [InlineData(FixedFormatType.Pdf, "pdf")]
    [InlineData(FixedFormatType.Xps, "xps")]
    public void ExportFixedFormat_CreatesRequestedFile(
        FixedFormatType formatType,
        string extension)
    {
        var testFile = _fixture.CreateTestFile();
        var outputPath = Path.Join(_fixture.TempDir, $"WorkbookExport_{Guid.NewGuid():N}.{extension}");
        WriteCell(testFile, "Printable workbook content");
        using var batch = ExcelSession.BeginBatch(testFile);

        var result = _commands.ExportFixedFormat(
            batch,
            outputPath,
            formatType,
            FixedFormatQuality.Standard,
            includeDocumentProperties: true,
            ignorePrintAreas: false,
            fromPage: null,
            toPage: null,
            openAfterPublish: false);

        Assert.True(result.Success);
        Assert.True(System.IO.File.Exists(outputPath));
        Assert.NotEmpty(System.IO.File.ReadAllBytes(outputPath));
        if (formatType == FixedFormatType.Pdf)
        {
            Assert.StartsWith(
                "%PDF",
                System.Text.Encoding.ASCII.GetString(System.IO.File.ReadAllBytes(outputPath), 0, 4));
        }
    }

    [Fact]
    public void ExternalLinks_ListUpdateAndBreak_RoundTrip()
    {
        var sourcePath = _fixture.CreateTestFile();
        var targetPath = _fixture.CreateTestFile();
        WriteCell(sourcePath, 10);
        WriteExternalFormula(targetPath, sourcePath);
        WriteCell(sourcePath, 42);

        using var batch = ExcelSession.BeginBatch(targetPath);
        var listResult = _commands.ListExternalLinks(batch);
        var link = Assert.Single(listResult.Links);

        Assert.Equal(Path.GetFullPath(sourcePath), link.Source, ignoreCase: true);

        var updateResult = _commands.UpdateExternalLink(batch, link.Source);
        var updatedValue = ReadCellValue(batch);
        var breakResult = _commands.BreakExternalLink(batch, link.Source);
        var linksAfterBreak = _commands.ListExternalLinks(batch);
        var formulaAfterBreak = ReadCellFormula(batch);

        Assert.True(updateResult.Success);
        Assert.Equal(42d, Convert.ToDouble(updatedValue, CultureInfo.InvariantCulture));
        Assert.True(breakResult.Success);
        Assert.Empty(linksAfterBreak.Links);
        Assert.Equal(42d, Convert.ToDouble(formulaAfterBreak, CultureInfo.InvariantCulture));
    }

    private static void WriteCell(string workbookPath, object value)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheet = (ExcelWorksheet)context.Book.Worksheets[1];
                cell = sheet.Range["A1"];
                cell.Value2 = value;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
        batch.Save();
    }

    private static void WriteExternalFormula(string workbookPath, string sourcePath)
    {
        var sourceDirectory = Path.GetDirectoryName(sourcePath)!.Replace("'", "''", StringComparison.Ordinal);
        var sourceFileName = Path.GetFileName(sourcePath);
        var formula = $"='{sourceDirectory}\\[{sourceFileName}]Sheet1'!$A$1";

        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheet = (ExcelWorksheet)context.Book.Worksheets[1];
                cell = sheet.Range["A1"];
                cell.Formula = formula;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
        batch.Save();
    }

    private static object? ReadCellValue(IExcelBatch batch)
    {
        return batch.Execute((context, _) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheet = (ExcelWorksheet)context.Book.Worksheets[1];
                cell = sheet.Range["A1"];
                return cell.Value2;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static object? ReadCellFormula(IExcelBatch batch)
    {
        return batch.Execute((context, _) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheet = (ExcelWorksheet)context.Book.Worksheets[1];
                cell = sheet.Range["A1"];
                return cell.Formula;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static bool IsWorkbookProtected(IExcelBatch batch)
    {
        return batch.Execute((ctx, ct) =>
            ctx.Book.ProtectStructure || ctx.Book.ProtectWindows);
    }
}
