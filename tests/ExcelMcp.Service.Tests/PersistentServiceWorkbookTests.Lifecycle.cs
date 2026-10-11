using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Xunit;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelSheets = Microsoft.Office.Interop.Excel.Sheets;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceWorkbookTests
{
    [Fact]
    public void SaveCopyAs_CreatesIndependentWorkbookCopy()
    {
        var copyPath = CreateOutputPath("WorkbookCopy", "xlsx");
        const string marker = "independent copy content";
        WriteCurrentWorkbookCell(marker);

        var result = _workbook.SaveCopyAs(
            _fixture.BatchToken,
            copyPath,
            overwrite: false);

        RequireSuccess(result);
        Assert.Equal(Path.GetFullPath(copyPath), result.FilePath, ignoreCase: true);
        Assert.True(File.Exists(copyPath));
        Assert.Equal(
            Path.GetFullPath(_fixture.WorkbookPath),
            GetContextWorkbookPath(),
            ignoreCase: true);
        WriteCurrentWorkbookCell("source changed after copy");
        Assert.Equal(marker, ReadSavedWorkbookCell(copyPath));
        Assert.Equal("source changed after copy", ReadCurrentWorkbookCellValue());
    }

    [Theory]
    [InlineData(WorkbookSaveFormat.Xlsx, "xlsx")]
    [InlineData(WorkbookSaveFormat.Xlsm, "xlsm")]
    [InlineData(WorkbookSaveFormat.Xlsb, "xlsb")]
    public void SaveAs_ChangesWorkbookFormatAndSessionPath(
        WorkbookSaveFormat format,
        string extension)
    {
        AssertSaveAsChangesWorkbookFormatAndSessionPath(format, extension);
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public void SaveAs_LegacyXls_ChangesWorkbookFormatAndSessionPath()
    {
        AssertSaveAsChangesWorkbookFormatAndSessionPath(WorkbookSaveFormat.Xls, "xls");
    }

    private void AssertSaveAsChangesWorkbookFormatAndSessionPath(
        WorkbookSaveFormat format,
        string extension)
    {
        var outputPath = CreateOutputPath("WorkbookSaveAs", extension);
        const string marker = "saved in the requested workbook format";
        WriteCurrentWorkbookCell(marker);
        RunWithRestoredWorkbookPath(() =>
        {
            var result = _workbook.SaveAs(
                _fixture.BatchToken,
                outputPath,
                format,
                overwrite: false);
            var info = RequireSuccess(_workbook.GetInfo(_fixture.BatchToken));

            RequireSuccess(result);
            Assert.Equal(Path.GetFullPath(outputPath), result.FilePath, ignoreCase: true);
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
            RestoreFixtureWorkbookPath();
            Assert.Equal(marker, ReadSavedWorkbookCell(outputPath));
        });
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

        RequireSuccess(result);
        Assert.Equal(Path.GetFullPath(outputPath), result.FilePath, ignoreCase: true);
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
        const string marker = "source retained after rejected save-as";
        WriteCurrentWorkbookCell(marker);
        var originalPath = GetContextWorkbookPath();

        var error = Assert.Throws<ArgumentException>(() =>
            _workbook.SaveAs(
                _fixture.BatchToken,
                outputPath,
                WorkbookSaveFormat.Xlsm,
                overwrite: true));

        Assert.Contains("extension", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(existingContent, File.ReadAllText(outputPath));
        Assert.Equal(originalPath, GetContextWorkbookPath());
        Assert.Equal(marker, ReadCurrentWorkbookCellValue());
    }

    [Fact]
    public void SaveCopyAs_InvalidExtensionWithOverwrite_PreservesExistingDestination()
    {
        var outputPath = CreateOutputPath("ExistingCopy", "xlsm");
        const string existingContent = "existing copy destination";
        File.WriteAllText(outputPath, existingContent);
        const string marker = "source retained after rejected copy";
        WriteCurrentWorkbookCell(marker);
        var originalPath = GetContextWorkbookPath();

        var error = Assert.Throws<ArgumentException>(() =>
            _workbook.SaveCopyAs(
                _fixture.BatchToken,
                outputPath,
                overwrite: true));

        Assert.Contains("extension", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(existingContent, File.ReadAllText(outputPath));
        Assert.Equal(originalPath, GetContextWorkbookPath());
        Assert.Equal(marker, ReadCurrentWorkbookCellValue());
    }

    [Fact]
    public void ExportFixedFormat_InvalidExtensionWithOverwrite_PreservesExistingDestination()
    {
        var outputPath = CreateOutputPath("ExistingExport", "xps");
        const string existingContent = "existing export destination";
        File.WriteAllText(outputPath, existingContent);
        const string marker = "source retained after rejected export";
        WriteCurrentWorkbookCell(marker);
        var originalPath = GetContextWorkbookPath();

        var error = Assert.Throws<ArgumentException>(() =>
            _workbook.ExportFixedFormat(
                _fixture.BatchToken,
                outputPath,
                FixedFormatType.Pdf,
                overwrite: true));

        Assert.Contains("extension", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(existingContent, File.ReadAllText(outputPath));
        Assert.Equal(originalPath, GetContextWorkbookPath());
        Assert.Equal(marker, ReadCurrentWorkbookCellValue());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Save_ExistingDestinationWithoutOverwrite_PreservesDestinationAndSession(bool copy)
    {
        var outputPath = _fixture.CreateBlankWorkbook(
            nameof(Save_ExistingDestinationWithoutOverwrite_PreservesDestinationAndSession), "Target");
        const string destinationMarker = "existing destination";
        WriteCell(outputPath, destinationMarker);
        const string sourceMarker = "retained source";
        WriteCurrentWorkbookCell(sourceMarker);
        var originalPath = GetContextWorkbookPath();

        var error = Assert.Throws<IOException>(() =>
        {
            if (copy)
            {
                _workbook.SaveCopyAs(_fixture.BatchToken, outputPath, overwrite: false);
            }
            else
            {
                _workbook.SaveAs(_fixture.BatchToken, outputPath, overwrite: false);
            }
        });

        Assert.Contains("already exists", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(originalPath, GetContextWorkbookPath());
        Assert.Equal(sourceMarker, ReadCurrentWorkbookCellValue());
        Assert.Equal(destinationMarker, ReadSavedWorkbookCell(outputPath));
        WriteCurrentWorkbookCell("usable after rejection");
        Assert.Equal("usable after rejection", ReadCurrentWorkbookCellValue());
    }

    [Fact]
    public void SaveAs_WithOverwrite_ReplacesExistingWorkbook()
    {
        var outputPath = _fixture.CreateBlankWorkbook(
            nameof(SaveAs_WithOverwrite_ReplacesExistingWorkbook),
            "Target");
        const string marker = "replacement workbook content";
        WriteCurrentWorkbookCell(marker);
        RunWithRestoredWorkbookPath(() =>
        {
            var result = _workbook.SaveAs(
                _fixture.BatchToken,
                outputPath,
                WorkbookSaveFormat.Xlsx,
                overwrite: true);

            RequireSuccess(result);
            Assert.Equal(Path.GetFullPath(outputPath), result.FilePath, ignoreCase: true);
            Assert.Equal(
                Path.GetFullPath(outputPath),
                GetContextWorkbookPath(),
                ignoreCase: true);
            Assert.Equal(
                Path.GetFullPath(outputPath),
                RequireSuccess(_workbook.GetInfo(_fixture.BatchToken)).FullName,
                ignoreCase: true);
            RestoreFixtureWorkbookPath();
            Assert.Equal(marker, ReadSavedWorkbookCell(outputPath));
        });
    }

    [Fact]
    public void SaveCopyAs_WithOverwrite_ReplacesExistingDestination()
    {
        var outputPath = CreateOutputPath("OverwriteCopy", "xlsx");
        File.WriteAllText(outputPath, "existing copy destination");
        const string marker = "replacement copy content";
        WriteCurrentWorkbookCell(marker);

        var result = _workbook.SaveCopyAs(
            _fixture.BatchToken,
            outputPath,
            overwrite: true);

        RequireSuccess(result);
        Assert.Equal(Path.GetFullPath(outputPath), result.FilePath, ignoreCase: true);
        Assert.Equal(Path.GetFullPath(_fixture.WorkbookPath), GetContextWorkbookPath(), ignoreCase: true);
        Assert.Equal(marker, ReadCurrentWorkbookCellValue());
        Assert.Equal(marker, ReadSavedWorkbookCell(outputPath));
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

        RequireSuccess(result);
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

        var listResult = RequireSuccess(_workbook.ListExternalLinks(_fixture.BatchToken));
        var link = Assert.Single(listResult.Links);
        Assert.Equal(Path.GetFullPath(sourcePath), link.Source, ignoreCase: true);

        var updateResult = _workbook.UpdateExternalLink(
            _fixture.BatchToken,
            link.Source);
        RequireSuccess(updateResult);
        var updatedValue = ReadCurrentWorkbookCellValue();
        Assert.Equal(42d, Convert.ToDouble(updatedValue, CultureInfo.InvariantCulture));
        var breakResult = _workbook.BreakExternalLink(
            _fixture.BatchToken,
            link.Source);
        RequireSuccess(breakResult);
        var linksAfterBreak = RequireSuccess(_workbook.ListExternalLinks(_fixture.BatchToken));
        var formulaAfterBreak = ReadCurrentWorkbookCellFormula();

        Assert.Empty(linksAfterBreak.Links);
        Assert.Equal(
            42d,
            Convert.ToDouble(formulaAfterBreak, CultureInfo.InvariantCulture));
    }

    private string CreateOutputPath(string prefix, string extension) =>
        Path.Combine(
            Path.GetDirectoryName(_fixture.WorkbookPath)!,
            $"{prefix}_{Guid.NewGuid():N}.{extension}");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExternalLinks_MissingSource_PreservesExistingLinkAndCalculation(bool breakLink)
    {
        var batch = _fixture.BatchToken;
        var source = _fixture.CreateBlankWorkbook(
            nameof(ExternalLinks_MissingSource_PreservesExistingLinkAndCalculation), "Source");
        WriteCell(source, 10);
        WriteExternalFormulaToCurrentWorkbook(source);
        var before = ReadCurrentWorkbookCellFormula();
        var missing = CreateOutputPath("MissingLink", "xlsx");
        var original = Assert.Single(RequireSuccess(_workbook.ListExternalLinks(batch)).Links);

        var error = Assert.Throws<InvalidOperationException>(() =>
        {
            if (breakLink)
            {
                _workbook.BreakExternalLink(batch, missing);
            }
            else
            {
                _workbook.UpdateExternalLink(batch, missing);
            }
        });

        Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
        var retained = Assert.Single(RequireSuccess(_workbook.ListExternalLinks(batch)).Links);
        Assert.Equal(original.Source, retained.Source);
        Assert.Equal(before, ReadCurrentWorkbookCellFormula());
        Assert.Equal(10d, Convert.ToDouble(ReadCurrentWorkbookCellValue(), CultureInfo.InvariantCulture));
        WriteCell(source, 42);
        RequireSuccess(_workbook.UpdateExternalLink(batch, original.Source));
        Assert.Equal(42d, Convert.ToDouble(ReadCurrentWorkbookCellValue(), CultureInfo.InvariantCulture));
    }

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
        RequireSuccess(result);
    }

    private void RunWithRestoredWorkbookPath(Action test) =>
        RunWithWorkbookCleanup(test, RestoreFixtureWorkbookPath);

    private void WriteCurrentWorkbookCell(object value) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            ExcelSheets? sheets = null;
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (ExcelWorksheet)sheets[1];
                cell = sheet.Range["A1"];
                cell.Value2 = value;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
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
            ExcelSheets? sheets = null;
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (ExcelWorksheet)sheets[1];
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
                ComUtilities.Release(ref sheets);
            }
        });

    private static void WriteCell(string workbookPath, object value)
    {
        using var batch = ExcelSession.BeginBatch(workbookPath);
        batch.Execute((context, _) =>
        {
            ExcelSheets? sheets = null;
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (ExcelWorksheet)sheets[1];
                cell = sheet.Range["A1"];
                cell.Value2 = value;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        batch.Save();
    }

    private object? ReadCurrentWorkbookCellValue() =>
        ReadCurrentWorkbookCell(readFormula: false);

    private static string? ReadSavedWorkbookCell(string path)
    {
        using var service = new ExcelMcpService();
        ServiceResponse Send(string command, object args, string? sessionId = null)
        {
            var response = service.ProcessAsync(new ServiceRequest
            {
                Command = command,
                SessionId = sessionId,
                Args = JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions),
                Source = "saved-workbook-verification"
            }).GetAwaiter().GetResult();
            Assert.True(response.Success, response.ErrorMessage);
            Assert.True(string.IsNullOrEmpty(response.ErrorMessage));
            return response;
        }

        var opened = Send("session.open", new { filePath = path });
        using var document = JsonDocument.Parse(opened.Result!);
        var sessionId = document.RootElement.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        Exception? failure = null;
        string? result = null;
        try
        {
            var read = Send("range.get-values",
                new { sheetName = "Sheet1", rangeAddress = "A1" }, sessionId);
            using var values = JsonDocument.Parse(read.Result!);
            var root = values.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean());
            Assert.True(!root.TryGetProperty("errorMessage", out var error)
                || error.ValueKind == JsonValueKind.Null
                || string.IsNullOrEmpty(error.GetString()));
            var cell = Assert.Single(Assert.Single(
                root.GetProperty("values").EnumerateArray()).EnumerateArray());
            Assert.Equal(JsonValueKind.String, cell.ValueKind);
            result = cell.GetString();
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            try
            {
                Send("session.close", new { save = false }, sessionId);
            }
            catch (Exception cleanupFailure)
            {
                failure = PersistentServiceCleanupFailures.Combine(failure, cleanupFailure);
            }
        }
        if (failure is not null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Throw(failure);
        }
        return result;
    }

    private object? ReadCurrentWorkbookCellFormula() =>
        ReadCurrentWorkbookCell(readFormula: true);

    private object? ReadCurrentWorkbookCell(bool readFormula) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            ExcelSheets? sheets = null;
            ExcelWorksheet? sheet = null;
            ExcelRange? cell = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (ExcelWorksheet)sheets[1];
                cell = sheet.Range["A1"];
                return readFormula ? cell.Formula : cell.Value2;
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
