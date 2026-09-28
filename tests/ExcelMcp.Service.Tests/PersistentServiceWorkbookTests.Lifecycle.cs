using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceWorkbookTests
{
    [Fact]
    public void SaveCopyAs_CreatesIndependentWorkbookCopy()
    {
        var copyPath = CreateOutputPath("WorkbookCopy", "xlsx");

        var result = _workbook.SaveCopyAs(
            _fixture.BatchToken,
            copyPath,
            overwrite: false);

        Assert.True(result.Success);
        Assert.True(File.Exists(copyPath));
        Assert.Equal(
            Path.GetFullPath(_fixture.WorkbookPath),
            GetContextWorkbookPath(),
            ignoreCase: true);
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
        var outputPath = CreateOutputPath("WorkbookSaveAs", extension);
        try
        {
            var result = _workbook.SaveAs(
                _fixture.BatchToken,
                outputPath,
                format,
                overwrite: false);
            var info = _workbook.GetInfo(_fixture.BatchToken);

            Assert.True(result.Success);
            Assert.True(File.Exists(outputPath));
            Assert.Equal(
                Path.GetFullPath(outputPath),
                GetContextWorkbookPath(),
                ignoreCase: true);
            Assert.Equal(
                Path.GetFullPath(outputPath),
                info.FullName,
                ignoreCase: true);
            Assert.Equal(extension, info.Format);
        }
        finally
        {
            RestoreFixtureWorkbookPath();
        }
    }

    [Theory]
    [InlineData(FixedFormatType.Pdf, "pdf")]
    [InlineData(FixedFormatType.Xps, "xps")]
    public void ExportFixedFormat_CreatesRequestedFile(
        FixedFormatType formatType,
        string extension)
    {
        var outputPath = CreateOutputPath("WorkbookExport", extension);
        WriteCurrentWorkbookCell("Printable workbook content");

        var result = _workbook.ExportFixedFormat(
            _fixture.BatchToken,
            outputPath,
            formatType,
            FixedFormatQuality.Standard,
            includeDocumentProperties: true,
            ignorePrintAreas: false,
            fromPage: null,
            toPage: null,
            openAfterPublish: false);

        Assert.True(result.Success);
        Assert.True(File.Exists(outputPath));
        Assert.NotEmpty(File.ReadAllBytes(outputPath));
        if (formatType == FixedFormatType.Pdf)
        {
            Assert.StartsWith(
                "%PDF",
                System.Text.Encoding.ASCII.GetString(
                    File.ReadAllBytes(outputPath),
                    0,
                    4));
        }
    }

    [Fact]
    public void SaveAs_InvalidFormatWithOverwrite_PreservesExistingDestination()
    {
        var outputPath = CreateOutputPath("ExistingSaveAs", "xlsx");
        const string existingContent = "existing save-as destination";
        File.WriteAllText(outputPath, existingContent);

        Assert.Throws<ArgumentException>(() =>
            _workbook.SaveAs(
                _fixture.BatchToken,
                outputPath,
                WorkbookSaveFormat.Xlsm,
                overwrite: true));

        Assert.Equal(existingContent, File.ReadAllText(outputPath));
    }

    [Fact]
    public void SaveCopyAs_InvalidExtensionWithOverwrite_PreservesExistingDestination()
    {
        var outputPath = CreateOutputPath("ExistingCopy", "xlsm");
        const string existingContent = "existing copy destination";
        File.WriteAllText(outputPath, existingContent);

        Assert.Throws<ArgumentException>(() =>
            _workbook.SaveCopyAs(
                _fixture.BatchToken,
                outputPath,
                overwrite: true));

        Assert.Equal(existingContent, File.ReadAllText(outputPath));
    }

    [Fact]
    public void ExportFixedFormat_InvalidExtensionWithOverwrite_PreservesExistingDestination()
    {
        var outputPath = CreateOutputPath("ExistingExport", "xps");
        const string existingContent = "existing export destination";
        File.WriteAllText(outputPath, existingContent);

        Assert.Throws<ArgumentException>(() =>
            _workbook.ExportFixedFormat(
                _fixture.BatchToken,
                outputPath,
                FixedFormatType.Pdf,
                overwrite: true));

        Assert.Equal(existingContent, File.ReadAllText(outputPath));
    }

    [Fact]
    public void SaveAs_WithOverwrite_ReplacesExistingWorkbook()
    {
        var outputPath = _fixture.CreateBlankWorkbook(
            nameof(SaveAs_WithOverwrite_ReplacesExistingWorkbook),
            "Target");
        try
        {
            var result = _workbook.SaveAs(
                _fixture.BatchToken,
                outputPath,
                WorkbookSaveFormat.Xlsx,
                overwrite: true);

            Assert.True(result.Success);
            Assert.Equal(
                Path.GetFullPath(outputPath),
                GetContextWorkbookPath(),
                ignoreCase: true);
            Assert.Equal(
                Path.GetFullPath(outputPath),
                _workbook.GetInfo(_fixture.BatchToken).FullName,
                ignoreCase: true);
        }
        finally
        {
            RestoreFixtureWorkbookPath();
        }
    }

    [Fact]
    public void SaveCopyAs_WithOverwrite_ReplacesExistingDestination()
    {
        var outputPath = CreateOutputPath("OverwriteCopy", "xlsx");
        File.WriteAllText(outputPath, "existing copy destination");

        var result = _workbook.SaveCopyAs(
            _fixture.BatchToken,
            outputPath,
            overwrite: true);

        Assert.True(result.Success);
        using var verifyBatch = ExcelSession.BeginBatch(outputPath);
        Assert.Equal(
            Path.GetFullPath(outputPath),
            verifyBatch.WorkbookPath,
            ignoreCase: true);
    }

    [Fact]
    public void ExportFixedFormat_WithOverwrite_ReplacesExistingDestination()
    {
        var outputPath = CreateOutputPath("OverwriteExport", "pdf");
        File.WriteAllText(outputPath, "existing export destination");
        WriteCurrentWorkbookCell("Printable workbook content");

        var result = _workbook.ExportFixedFormat(
            _fixture.BatchToken,
            outputPath,
            FixedFormatType.Pdf,
            overwrite: true);

        Assert.True(result.Success);
        Assert.StartsWith(
            "%PDF",
            System.Text.Encoding.ASCII.GetString(
                File.ReadAllBytes(outputPath),
                0,
                4));
    }

    [Fact]
    public void ExternalLinks_ListUpdateAndBreak_RoundTrip()
    {
        var sourcePath = _fixture.CreateBlankWorkbook(
            nameof(ExternalLinks_ListUpdateAndBreak_RoundTrip),
            "Source");
        WriteCell(sourcePath, 10);
        WriteExternalFormulaToCurrentWorkbook(sourcePath);
        WriteCell(sourcePath, 42);

        var listResult = _workbook.ListExternalLinks(_fixture.BatchToken);
        var link = Assert.Single(listResult.Links);
        Assert.Equal(Path.GetFullPath(sourcePath), link.Source, ignoreCase: true);

        var updateResult = _workbook.UpdateExternalLink(
            _fixture.BatchToken,
            link.Source);
        var updatedValue = ReadCurrentWorkbookCellValue();
        var breakResult = _workbook.BreakExternalLink(
            _fixture.BatchToken,
            link.Source);
        var linksAfterBreak = _workbook.ListExternalLinks(_fixture.BatchToken);
        var formulaAfterBreak = ReadCurrentWorkbookCellFormula();

        Assert.True(updateResult.Success);
        Assert.Equal(
            42d,
            Convert.ToDouble(updatedValue, CultureInfo.InvariantCulture));
        Assert.True(breakResult.Success);
        Assert.Empty(linksAfterBreak.Links);
        Assert.Equal(
            42d,
            Convert.ToDouble(formulaAfterBreak, CultureInfo.InvariantCulture));
    }

    private string CreateOutputPath(string prefix, string extension) =>
        Path.Combine(
            Path.GetDirectoryName(_fixture.WorkbookPath)!,
            $"{prefix}_{Guid.NewGuid():N}.{extension}");

    private string GetContextWorkbookPath() =>
        _fixture.ExecuteRawVerification((context, _) => context.WorkbookPath);

    private void RestoreFixtureWorkbookPath()
    {
        if (string.Equals(
                GetContextWorkbookPath(),
                Path.GetFullPath(_fixture.WorkbookPath),
                StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        var result = _workbook.SaveAs(
            _fixture.BatchToken,
            _fixture.WorkbookPath,
            WorkbookSaveFormat.Xlsx,
            overwrite: true);
        Assert.True(result.Success, result.ErrorMessage);
    }

    private void WriteCurrentWorkbookCell(object value) =>
        _fixture.ExecuteRawVerification((context, _) =>
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

    private void WriteExternalFormulaToCurrentWorkbook(string sourcePath)
    {
        var sourceDirectory = Path.GetDirectoryName(sourcePath)!
            .Replace("'", "''", StringComparison.Ordinal);
        var sourceFileName = Path.GetFileName(sourcePath);
        var formula = $"='{sourceDirectory}\\[{sourceFileName}]Sheet1'!$A$1";
        WriteCurrentWorkbookCell(formula, setFormula: true);
    }

    private void WriteCurrentWorkbookCell(object value, bool setFormula) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheet = (ExcelWorksheet)context.Book.Worksheets[1];
                cell = sheet.Range["A1"];
                if (setFormula)
                {
                    cell.Formula = value;
                }
                else
                {
                    cell.Value2 = value;
                }
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

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

    private object? ReadCurrentWorkbookCellValue() =>
        ReadCurrentWorkbookCell(readFormula: false);

    private object? ReadCurrentWorkbookCellFormula() =>
        ReadCurrentWorkbookCell(readFormula: true);

    private object? ReadCurrentWorkbookCell(bool readFormula) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheet = (ExcelWorksheet)context.Book.Worksheets[1];
                cell = sheet.Range["A1"];
                return readFormula ? cell.Formula : cell.Value2;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
}
