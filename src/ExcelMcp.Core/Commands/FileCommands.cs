using Sbroenne.ExcelMcp.Core.Models;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

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
            if (extension != ".xlsx" && extension != ".xlsm")
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = "File must have .xlsx or .xlsm extension",
                    FilePath = filePath,
                    Action = "create-empty"
                };
            }
            
            // Check if file already exists
            if (File.Exists(filePath) && !overwriteIfExists)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = $"File already exists: {filePath}",
                    FilePath = filePath,
                    Action = "create-empty"
                };
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
                    return new OperationResult
                    {
                        Success = false,
                        ErrorMessage = $"Failed to create directory: {ex.Message}",
                        FilePath = filePath,
                        Action = "create-empty"
                    };
                }
            }

            // Create Excel workbook with COM automation
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                return new OperationResult
                {
                    Success = false,
                    ErrorMessage = "Excel is not installed. Cannot create Excel files.",
                    FilePath = filePath,
                    Action = "create-empty"
                };
            }

#pragma warning disable IL2072 // COM interop is not AOT compatible
            dynamic excel = Activator.CreateInstance(excelType)!;
#pragma warning restore IL2072
            try
            {
                excel.Visible = false;
                excel.DisplayAlerts = false;
                
                // Create new workbook
                dynamic workbook = excel.Workbooks.Add();
                
                // Optional: Set up a basic structure
                dynamic sheet = workbook.Worksheets.Item(1);
                sheet.Name = "Sheet1";
                
                // Add a comment to indicate this was created by ExcelCLI
                sheet.Range["A1"].AddComment($"Created by ExcelCLI on {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sheet.Range["A1"].Comment.Visible = false;
                
                // Save the workbook with appropriate format
                if (extension == ".xlsm")
                {
                    // Save as macro-enabled workbook (format 52)
                    workbook.SaveAs(filePath, 52);
                }
                else
                {
                    // Save as regular workbook (format 51)
                    workbook.SaveAs(filePath, 51);
                }
                
                workbook.Close(false);
                
                return new OperationResult
                {
                    Success = true,
                    FilePath = filePath,
                    Action = "create-empty"
                };
            }
            finally
            {
                try { excel.Quit(); } catch { }
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); } catch { }
                
                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                // Small delay for Excel to fully close
                System.Threading.Thread.Sleep(100);
            }
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
    
    /// <inheritdoc />
    public FileValidationResult Validate(string filePath)
    {
        try
        {
            filePath = Path.GetFullPath(filePath);
            
            var result = new FileValidationResult
            {
                Success = true,
                FilePath = filePath,
                Exists = File.Exists(filePath)
            };
            
            if (result.Exists)
            {
                var fileInfo = new FileInfo(filePath);
                result.Size = fileInfo.Length;
                result.Extension = fileInfo.Extension;
                result.LastModified = fileInfo.LastWriteTime;
                result.IsValid = result.Extension.ToLowerInvariant() == ".xlsx" || 
                                 result.Extension.ToLowerInvariant() == ".xlsm";
            }
            else
            {
                result.Extension = Path.GetExtension(filePath);
                result.IsValid = false;
            }
            
            return result;
        }
        catch (Exception ex)
        {
            return new FileValidationResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = filePath,
                Exists = false,
                IsValid = false
            };
        }
    }
}
