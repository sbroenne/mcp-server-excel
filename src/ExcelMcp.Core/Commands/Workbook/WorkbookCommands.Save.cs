using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

public partial class WorkbookCommands
{
    internal static Action<string>? BeforeOutputCommitHookForTests { get; set; }

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

        var owner = batch switch
        {
            ExcelBatch owned => owned,
            AttachedWorkbookBatch attached => attached.Owner,
            _ => null
        };
        if (owner is not null)
            ExcelBusyException.ThrowIfNotReady(owner.GetExecutionPreflightState(), "save-as", batch as IExcelBatchWindowControl);

        return batch.Execute((context, _) =>
        {
            ExcelBusyException.ThrowIfNotReady(
                ExcelDialogProbe.Read(context.ProcessIdentity, batch.Logger), "save-as", batch as IExcelBatchWindowControl);
            ExcelBusyException.ThrowIfNotReady(
                WorkbookRefreshProbe.Read(context.App, context.Book), "save-as");
            var originalFullName = context.Book.FullName;
            var displayAlerts = context.App.DisplayAlerts;
            var targetAccepted = false;
            Exception? operationFailure = null;
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
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                OperationCleanup.Run(() =>
                {
                    context.App.DisplayAlerts = displayAlerts;
                    // OneDrive can resolve a local target to a SharePoint URL.
                    targetAccepted =
                        string.Equals(context.Book.Name, Path.GetFileName(normalizedPath), StringComparison.OrdinalIgnoreCase) &&
                        (string.Equals(batch.WorkbookPath, normalizedPath, StringComparison.OrdinalIgnoreCase) ||
                         !string.Equals(context.Book.FullName, originalFullName, StringComparison.OrdinalIgnoreCase));
                    if (targetAccepted)
                        batch.UpdateWorkbookPath(normalizedPath);
                }, operationFailure);
            }

            if (!context.Book.Saved || !targetAccepted)
            {
                throw new InvalidOperationException(
                    "Workbook was not saved to the requested path. Excel cancelled Save As or left unsaved changes. " +
                    "The workbook remains open; inspect its saved state and fullName before continuing.");
            }

            return new OperationResult
            {
                Success = true,
                Action = "save-as",
                FilePath = normalizedPath,
                Message = $"Workbook saved as '{normalizedPath}'"
            };
        });
    }

    /// <inheritdoc />
    public OperationResult SaveCopyAs(IExcelBatch batch, string targetPath, bool overwrite = false)
    {
        var normalizedPath = ValidateOutputPath(targetPath, overwrite);
        var outputExtension = Path.GetExtension(normalizedPath);
        var writePath = GetWritePath(normalizedPath, overwrite);
        return batch.Execute((context, _) =>
        {
            var currentExtension = string.IsNullOrEmpty(context.Book.Path)
                ? context.Book.FileFormat switch
                {
                    Excel.XlFileFormat.xlOpenXMLWorkbook => ".xlsx",
                    Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled => ".xlsm",
                    Excel.XlFileFormat.xlExcel12 => ".xlsb",
                    Excel.XlFileFormat.xlExcel8 or Excel.XlFileFormat.xlWorkbookNormal => ".xls",
                    var format => throw new InvalidOperationException(
                        $"Save Copy As cannot determine an output extension for never-saved workbook format '{format}'.")
                }
                : Path.GetExtension(context.WorkbookPath);
            if (!string.Equals(currentExtension, outputExtension, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    $"Save Copy As preserves the current workbook format. Output extension must be '{currentExtension}'.",
                    nameof(targetPath));
            }
            Exception? operationFailure = null;
            try
            {
                context.Book.SaveCopyAs(writePath);
                BeforeOutputCommitHookForTests?.Invoke(writePath);
                CommitOutput(writePath, normalizedPath);
                return new OperationResult
                {
                    Success = true,
                    Action = "save-copy-as",
                    FilePath = normalizedPath,
                    Message = $"Workbook copy saved to '{normalizedPath}'"
                };
            }
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                OperationCleanup.Run(() => DeleteTemporaryOutput(writePath, normalizedPath), operationFailure);
            }
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

        var writePath = GetWritePath(normalizedPath, overwrite);
        var openAfterCommit =
            openAfterPublish &&
            !string.Equals(writePath, normalizedPath, StringComparison.OrdinalIgnoreCase);
        return batch.Execute((context, _) =>
        {
            Exception? operationFailure = null;
            try
            {
                context.Book.ExportAsFixedFormat(
                    formatType == FixedFormatType.Pdf
                        ? Excel.XlFixedFormatType.xlTypePDF
                        : Excel.XlFixedFormatType.xlTypeXPS,
                    writePath,
                    quality == FixedFormatQuality.Standard
                        ? Excel.XlFixedFormatQuality.xlQualityStandard
                        : Excel.XlFixedFormatQuality.xlQualityMinimum,
                    includeDocumentProperties,
                    ignorePrintAreas,
                    fromPage ?? Type.Missing,
                    toPage ?? Type.Missing,
                    openAfterPublish && !openAfterCommit,
                    Type.Missing);

                BeforeOutputCommitHookForTests?.Invoke(writePath);
                CommitOutput(writePath, normalizedPath);
                if (openAfterCommit)
                {
                    using var process = System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo(normalizedPath)
                        {
                            UseShellExecute = true
                        });
                }

                return new OperationResult
                {
                    Success = true,
                    Action = "export-fixed-format",
                    FilePath = normalizedPath,
                    Message = $"Workbook exported to '{normalizedPath}'"
                };
            }
            catch (Exception ex)
            {
                operationFailure = ex;
                throw;
            }
            finally
            {
                OperationCleanup.Run(() => DeleteTemporaryOutput(writePath, normalizedPath), operationFailure);
            }
        });
    }

    internal static string ValidateOutputPath(string outputPath, bool overwrite)
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
        }

        return normalizedPath;
    }

    internal static string GetWritePath(string normalizedPath, bool overwrite)
    {
        if (!overwrite || !File.Exists(normalizedPath))
        {
            return normalizedPath;
        }

        var directory = Path.GetDirectoryName(normalizedPath)!;
        var fileName = Path.GetFileNameWithoutExtension(normalizedPath);
        var extension = Path.GetExtension(normalizedPath);
        return Path.Combine(directory, $".{fileName}.{Guid.NewGuid():N}.tmp{extension}");
    }

    internal static void CommitOutput(string writePath, string normalizedPath)
    {
        if (!string.Equals(writePath, normalizedPath, StringComparison.OrdinalIgnoreCase))
        {
            File.Move(writePath, normalizedPath, overwrite: true);
        }
    }

    internal static void DeleteTemporaryOutput(string writePath, string normalizedPath)
    {
        if (!string.Equals(writePath, normalizedPath, StringComparison.OrdinalIgnoreCase) &&
            File.Exists(writePath))
        {
            File.Delete(writePath);
        }
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
