using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    /// <inheritdoc />
    public OperationResult SaveAs(
        IExcelBatch batch,
        string targetPath,
        WorkbookSaveFormat format = WorkbookSaveFormat.Auto,
        bool overwrite = false)
    {
        var normalizedPath = ValidateOutputPath(targetPath, overwrite);
        var resolvedFormat = ResolveSaveFormat(normalizedPath, format);
        ValidateSaveExtension(normalizedPath, resolvedFormat);

        var result = batch.Execute((context, _) =>
        {
            var displayAlerts = context.App.DisplayAlerts;
            try
            {
                context.App.DisplayAlerts = false;
                context.Book.SaveAs(
                    normalizedPath,
                    ToExcelFileFormat(resolvedFormat),
                    Type.Missing,
                    Type.Missing,
                    false,
                    false,
                    Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                    false,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);
            }
            finally
            {
                context.App.DisplayAlerts = displayAlerts;
            }

            return new OperationResult
            {
                Success = true,
                Action = "save-as",
                FilePath = normalizedPath,
                Message = $"Workbook saved as '{normalizedPath}'"
            };
        });

        batch.UpdateWorkbookPath(normalizedPath);
        return result;
    }

    /// <inheritdoc />
    public OperationResult SaveCopyAs(IExcelBatch batch, string targetPath, bool overwrite = false)
    {
        var normalizedPath = ValidateOutputPath(targetPath, overwrite);
        var currentExtension = Path.GetExtension(batch.WorkbookPath);
        var outputExtension = Path.GetExtension(normalizedPath);
        if (!string.Equals(currentExtension, outputExtension, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"Save Copy As preserves the current workbook format. Output extension must be '{currentExtension}'.",
                nameof(targetPath));
        }

        return batch.Execute((context, _) =>
        {
            context.Book.SaveCopyAs(normalizedPath);
            return new OperationResult
            {
                Success = true,
                Action = "save-copy-as",
                FilePath = normalizedPath,
                Message = $"Workbook copy saved to '{normalizedPath}'"
            };
        });
    }

    /// <inheritdoc />
    public OperationResult ExportFixedFormat(
        IExcelBatch batch,
        string targetPath,
        FixedFormatType formatType = FixedFormatType.Pdf,
        FixedFormatQuality quality = FixedFormatQuality.Standard,
        bool includeDocumentProperties = true,
        bool ignorePrintAreas = false,
        int? fromPage = null,
        int? toPage = null,
        bool openAfterPublish = false,
        bool overwrite = false)
    {
        ValidatePageRange(fromPage, toPage);
        var normalizedPath = ValidateOutputPath(targetPath, overwrite);
        var expectedExtension = formatType == FixedFormatType.Pdf ? ".pdf" : ".xps";
        if (!string.Equals(Path.GetExtension(normalizedPath), expectedExtension, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"{formatType} export requires the '{expectedExtension}' file extension.",
                nameof(targetPath));
        }

        return batch.Execute((context, _) =>
        {
            context.Book.ExportAsFixedFormat(
                formatType == FixedFormatType.Pdf
                    ? Excel.XlFixedFormatType.xlTypePDF
                    : Excel.XlFixedFormatType.xlTypeXPS,
                normalizedPath,
                quality == FixedFormatQuality.Standard
                    ? Excel.XlFixedFormatQuality.xlQualityStandard
                    : Excel.XlFixedFormatQuality.xlQualityMinimum,
                includeDocumentProperties,
                ignorePrintAreas,
                fromPage ?? Type.Missing,
                toPage ?? Type.Missing,
                openAfterPublish,
                Type.Missing);

            return new OperationResult
            {
                Success = true,
                Action = "export-fixed-format",
                FilePath = normalizedPath,
                Message = $"Workbook exported to '{normalizedPath}'"
            };
        });
    }

    private static string ValidateOutputPath(string outputPath, bool overwrite)
    {
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new ArgumentException("Output path cannot be empty.", nameof(outputPath));
        }

        var normalizedPath = Path.GetFullPath(outputPath);
        var directory = Path.GetDirectoryName(normalizedPath);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException($"Output directory does not exist: '{directory}'.");
        }

        if (File.Exists(normalizedPath))
        {
            if (!overwrite)
            {
                throw new IOException($"Output file already exists: '{normalizedPath}'.");
            }

            File.Delete(normalizedPath);
        }

        return normalizedPath;
    }

    private static WorkbookSaveFormat ResolveSaveFormat(string outputPath, WorkbookSaveFormat format)
    {
        if (format != WorkbookSaveFormat.Auto)
        {
            return format;
        }

        return Path.GetExtension(outputPath).ToLowerInvariant() switch
        {
            ".xlsx" => WorkbookSaveFormat.Xlsx,
            ".xlsm" => WorkbookSaveFormat.Xlsm,
            ".xlsb" => WorkbookSaveFormat.Xlsb,
            ".xls" => WorkbookSaveFormat.Xls,
            var extension => throw new ArgumentException(
                $"Cannot infer workbook format from extension '{extension}'. Supported extensions: .xlsx, .xlsm, .xlsb, .xls.")
        };
    }

    private static void ValidateSaveExtension(string outputPath, WorkbookSaveFormat format)
    {
        var expectedExtension = format switch
        {
            WorkbookSaveFormat.Xlsx => ".xlsx",
            WorkbookSaveFormat.Xlsm => ".xlsm",
            WorkbookSaveFormat.Xlsb => ".xlsb",
            WorkbookSaveFormat.Xls => ".xls",
            _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported Save As format.")
        };

        if (!string.Equals(Path.GetExtension(outputPath), expectedExtension, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"Save As format '{format}' requires the '{expectedExtension}' file extension.",
                nameof(outputPath));
        }
    }

    private static Excel.XlFileFormat ToExcelFileFormat(WorkbookSaveFormat format)
    {
        return format switch
        {
            WorkbookSaveFormat.Xlsx => Excel.XlFileFormat.xlOpenXMLWorkbook,
            WorkbookSaveFormat.Xlsm => Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled,
            WorkbookSaveFormat.Xlsb => Excel.XlFileFormat.xlExcel12,
            WorkbookSaveFormat.Xls => Excel.XlFileFormat.xlExcel8,
            _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported Save As format.")
        };
    }

    private static void ValidatePageRange(int? fromPage, int? toPage)
    {
        if (fromPage is <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(fromPage), "fromPage must be greater than zero.");
        }

        if (toPage is <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(toPage), "toPage must be greater than zero.");
        }

        if (fromPage.HasValue && toPage.HasValue && fromPage > toPage)
        {
            throw new ArgumentException("fromPage cannot be greater than toPage.");
        }
    }
}
