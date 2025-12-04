using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands implementation
/// </summary>
public class FileCommands : IFileCommands
{
    /// <inheritdoc />
    public void CreateEmpty(string filePath, bool overwriteIfExists = false)
    {
        filePath = Path.GetFullPath(filePath);

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension is not ".xlsx" and not ".xlsm")
        {
            throw new ArgumentException("File must have .xlsx or .xlsm extension", nameof(filePath));
        }

        if (File.Exists(filePath) && !overwriteIfExists)
        {
            throw new InvalidOperationException($"File already exists: {filePath}. Use overwriteIfExists=true to overwrite.");
        }

        string? directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            try
            {
                Directory.CreateDirectory(directory);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to create directory '{directory}': {ex.Message}", ex);
            }
        }

        bool isMacroEnabled = extension == ".xlsm";
        CreateNewWorkbookOnStaThread(filePath, isMacroEnabled);
    }

    /// <summary>
    /// Creates a new Excel workbook directly on an STA thread without using batch API.
    /// This is faster and avoids session disposal overhead for simple file creation.
    /// </summary>
    private static void CreateNewWorkbookOnStaThread(string filePath, bool isMacroEnabled)
    {
        var completion = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);

        var thread = new Thread(() =>
        {
            dynamic? excel = null;
            dynamic? workbook = null;

            try
            {
                OleMessageFilter.Register();

                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel is not installed or not properly registered.");
                }

#pragma warning disable IL2072
                excel = Activator.CreateInstance(excelType);
#pragma warning restore IL2072

                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Use US-style separators (. for decimal, , for thousands) regardless of system locale
                // This ensures format codes like "$#,##0.00" work consistently across all locales
                excel.UseSystemSeparators = false;
                excel.DecimalSeparator = ".";
                excel.ThousandsSeparator = ",";

                workbook = excel.Workbooks.Add();

                // Save the workbook with 5-minute timeout (Excel automatically creates Sheet1)
                var saveAsTask = Task.Run(() =>
                {
                    if (isMacroEnabled)
                    {
                        workbook.SaveAs(filePath, 52); // xlOpenXMLWorkbookMacroEnabled
                    }
                    else
                    {
                        workbook.SaveAs(filePath, 51); // xlOpenXMLWorkbook
                    }
                });

                using var saveCts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
                if (!saveAsTask.Wait(TimeSpan.FromMinutes(5), saveCts.Token))
                {
                    throw new TimeoutException(
                        $"SaveAs operation for '{Path.GetFileName(filePath)}' exceeded 5 minutes. " +
                        "Check disk performance and antivirus settings.");
                }

                completion.TrySetResult(null);
            }
            catch (Exception ex)
            {
                completion.TrySetException(new InvalidOperationException($"Failed to create Excel file: {ex.Message}", ex));
            }
            finally
            {
                // Use ExcelShutdownService for resilient close and quit
                // save=false: file was already saved via SaveAs
                if (workbook != null || excel != null)
                {
                    ExcelShutdownService.CloseAndQuit(workbook, excel, false, filePath, null);
                }

                OleMessageFilter.Revoke();
            }
        })
        {
            IsBackground = true,
            Name = $"ExcelCreate-{Path.GetFileName(filePath)}"
        };

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();

        try
        {
            completion.Task.GetAwaiter().GetResult();
        }
        finally
        {
            thread.Join();
        }
    }

    /// <inheritdoc />
    public FileValidationInfo Test(string filePath)
    {
        filePath = Path.GetFullPath(filePath);

        bool exists = File.Exists(filePath);
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        bool isValidExtension = extension is ".xlsx" or ".xlsm";

        long size = 0;
        DateTime lastModified = DateTime.MinValue;

        if (exists)
        {
            var fileInfo = new FileInfo(filePath);
            size = fileInfo.Length;
            lastModified = fileInfo.LastWriteTime;
        }

        string? message = !exists
            ? $"File not found: {filePath}"
            : !isValidExtension ? $"Invalid file extension. Expected .xlsx or .xlsm, got {extension}" : null;

        return new FileValidationInfo
        {
            FilePath = filePath,
            Exists = exists,
            Size = size,
            Extension = extension,
            LastModified = lastModified,
            IsValid = exists && isValidExtension,
            Message = message
        };
    }

}


