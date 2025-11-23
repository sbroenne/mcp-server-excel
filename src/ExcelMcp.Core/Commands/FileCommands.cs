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
    public OperationResult CreateEmpty(string filePath, bool overwriteIfExists = false)
    {
        try
        {
            filePath = Path.GetFullPath(filePath);

            // Validate file extension
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension is not ".xlsx" and not ".xlsm")
            {
                throw new ArgumentException("File must have .xlsx or .xlsm extension", nameof(filePath));
            }

            // Check if file already exists
            if (File.Exists(filePath) && !overwriteIfExists)
            {
                throw new ArgumentException($"File already exists: {filePath}. Use overwriteIfExists=true to overwrite.", nameof(filePath));
            }

            // Ensure directory exists
            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                try
                {
                    Directory.CreateDirectory(directory);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to create directory: {ex.Message}", ex);
                }
            }

            // Create Excel workbook directly on STA thread - no batch session needed
            bool isMacroEnabled = extension == ".xlsm";

            return CreateNewWorkbookOnStaThread(filePath, isMacroEnabled);
        }
        catch (Exception ex)
        {
            return new OperationResult
            {
                Success = false,
                ErrorMessage = $"Failed to create Excel file: {ex.Message}",
                FilePath = filePath,
                Action = "create-empty"
            };
        }
    }

    /// <summary>
    /// Creates a new Excel workbook directly on an STA thread without using batch API.
    /// This is faster and avoids session disposal overhead for simple file creation.
    /// </summary>
    private static OperationResult CreateNewWorkbookOnStaThread(string filePath, bool isMacroEnabled)
    {
        var completion = new TaskCompletionSource<OperationResult>(TaskCreationOptions.RunContinuationsAsynchronously);

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

                completion.SetResult(new OperationResult
                {
                    Success = true,
                    FilePath = filePath,
                    Action = "create-empty"
                });
            }
            catch (Exception ex)
            {
                completion.SetResult(new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to create Excel file: {ex.Message}",
                    FilePath = filePath,
                    Action = "create-empty"
                });
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
        thread.Join(); // Wait for thread to complete

        return completion.Task.Result;
    }

    /// <inheritdoc />
    public FileValidationResult Test(string filePath)
    {
        try
        {
            filePath = Path.GetFullPath(filePath);

            // Check if file exists
            bool exists = File.Exists(filePath);

            // Get file extension
            string extension = exists ? Path.GetExtension(filePath).ToLowerInvariant() : "";

            // Validate extension
            bool isValidExtension = extension is ".xlsx" or ".xlsm";

            // Get file info if exists
            long size = 0;
            DateTime lastModified = DateTime.MinValue;

            if (exists)
            {
                var fileInfo = new FileInfo(filePath);
                size = fileInfo.Length;
                lastModified = fileInfo.LastWriteTime;
            }

            return new FileValidationResult
            {
                Success = exists && isValidExtension,
                ErrorMessage = !exists ? $"File not found: {filePath}"
                    : !isValidExtension ? $"Invalid file extension. Expected .xlsx or .xlsm, got {extension}"
                    : null,
                FilePath = filePath,
                Exists = exists,
                Size = size,
                Extension = extension,
                LastModified = lastModified,
                IsValid = exists && isValidExtension
            };
        }
        catch (Exception ex)
        {
            return new FileValidationResult
            {
                Success = false,
                ErrorMessage = $"Failed to validate file: {ex.Message}",
                FilePath = filePath,
                Exists = false,
                Size = 0,
                Extension = "",
                LastModified = DateTime.MinValue,
                IsValid = false
            };
        }
    }

}


