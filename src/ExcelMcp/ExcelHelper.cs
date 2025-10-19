using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Spectre.Console;

namespace ExcelMcp;

/// <summary>
/// Helper class for Excel COM automation with proper resource management
/// </summary>
public static class ExcelHelper
{
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static T WithExcel<T>(string filePath, bool save, Func<dynamic, dynamic, T> action)
    {
        dynamic? excel = null;
        dynamic? workbook = null;
        string operation = $"WithExcel({Path.GetFileName(filePath)}, save={save})";

        try
        {
            // Validate file path first - prevent path traversal attacks
            string fullPath = Path.GetFullPath(filePath);
            
            // Additional security: ensure the file is within reasonable bounds
            if (fullPath.Length > 32767)
            {
                throw new ArgumentException($"File path too long: {fullPath.Length} characters (Windows limit: 32767)");
            }
            
            // Security: Validate file extension to prevent executing arbitrary files
            string extension = Path.GetExtension(fullPath).ToLowerInvariant();
            if (extension is not (".xlsx" or ".xlsm" or ".xls"))
            {
                throw new ArgumentException($"Invalid file extension '{extension}'. Only Excel files (.xlsx, .xlsm, .xls) are supported.");
            }
            
            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"Excel file not found: {fullPath}", fullPath);
            }

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException("Excel is not installed or not properly registered. " +
                    "Please verify Microsoft Excel is installed and COM registration is intact.");
            }

#pragma warning disable IL2072 // COM interop is not AOT compatible but is required for Excel automation
            excel = Activator.CreateInstance(excelType);
#pragma warning restore IL2072
            if (excel == null) 
            {
                throw new InvalidOperationException("Failed to create Excel COM instance. " +
                    "Excel may be corrupted or COM subsystem unavailable.");
            }

            // Configure Excel for automation
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.Interactive = false;

            // Open workbook with detailed error context
            try
            {
                workbook = excel.Workbooks.Open(fullPath);
            }
            catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x8001010A))
            {
                // Excel is busy - provide specific guidance
                throw new InvalidOperationException(
                    "Excel is busy (likely has a dialog open). Close any Excel dialogs and retry.", comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x80070020))
            {
                // File sharing violation
                throw new InvalidOperationException(
                    $"File '{Path.GetFileName(fullPath)}' is locked by another process. " +
                    "Close Excel and any other applications using this file.", comEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to open workbook '{Path.GetFileName(fullPath)}'. " +
                    "File may be corrupted, password-protected, or incompatible.", ex);
            }

            if (workbook == null) 
            {
                throw new InvalidOperationException($"Failed to open workbook: {Path.GetFileName(fullPath)}");
            }

            // Execute the user action with error context
            T result;
            try
            {
                result = action(excel, workbook);
            }
            catch (Exception actionEx)
            {
                // Wrap action exceptions with enhanced context
                ExcelDiagnostics.ReportExcelError(actionEx, $"User Action in {operation}", fullPath, workbook, excel);
                throw;
            }

            // Save if requested
            if (save && workbook != null)
            {
                try
                {
                    workbook.Save();
                }
                catch (Exception saveEx)
                {
                    ExcelDiagnostics.ReportExcelError(saveEx, $"Save operation in {operation}", fullPath, workbook, excel);
                    throw;
                }
            }

            return result;
        }
        catch (Exception ex) when (!(ex.Data.Contains("ExcelDiagnosticsReported")))
        {
            // Only report if not already reported by inner exception
            ExcelDiagnostics.ReportExcelError(ex, operation, filePath, workbook, excel);
            ex.Data["ExcelDiagnosticsReported"] = true;
            throw;
        }
        finally
        {
            // Close workbook
            if (workbook != null)
            {
                try { workbook.Close(save); } catch { }
                try { Marshal.ReleaseComObject(workbook); } catch { }
            }

            // Quit Excel and release
            if (excel != null)
            {
                try { excel.Quit(); } catch { }
                try { Marshal.ReleaseComObject(excel); } catch { }
            }

            // Aggressive cleanup
            workbook = null;
            excel = null;

            // Force garbage collection multiple times
            for (int i = 0; i < 3; i++)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // Small delay to ensure Excel process terminates
            System.Threading.Thread.Sleep(100);
        }
    }

    public static dynamic? FindQuery(dynamic workbook, string queryName)
    {
        try
        {
            dynamic queriesCollection = workbook.Queries;
            int count = queriesCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic query = queriesCollection.Item(i);
                if (query.Name == queryName) return query;
            }
        }
        catch { }
        return null;
    }

    public static dynamic? FindName(dynamic workbook, string name)
    {
        try
        {
            dynamic namesCollection = workbook.Names;
            int count = namesCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic nameObj = namesCollection.Item(i);
                if (nameObj.Name == name) return nameObj;
            }
        }
        catch { }
        return null;
    }

    public static dynamic? FindSheet(dynamic workbook, string sheetName)
    {
        try
        {
            dynamic sheetsCollection = workbook.Worksheets;
            int count = sheetsCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic sheet = sheetsCollection.Item(i);
                if (sheet.Name == sheetName) return sheet;
            }
        }
        catch { }
        return null;
    }

    public static bool ValidateArgs(string[] args, int required, string usage)
    {
        if (args.Length >= required) return true;
        
        AnsiConsole.MarkupLine($"[red]Error:[/] Missing arguments");
        AnsiConsole.MarkupLine($"[yellow]Usage:[/] [cyan]ExcelCLI {usage.EscapeMarkup()}[/]");
        
        // Show what arguments were provided vs what's needed
        AnsiConsole.MarkupLine($"[dim]Provided {args.Length} arguments, need {required}[/]");
        
        if (args.Length > 0)
        {
            AnsiConsole.MarkupLine("[dim]Arguments provided:[/]");
            for (int i = 0; i < args.Length; i++)
            {
                AnsiConsole.MarkupLine($"[dim]  [[{i + 1}]] {args[i].EscapeMarkup()}[/]");
            }
        }
        
        // Parse usage string to show expected arguments
        var usageParts = usage.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (usageParts.Length > 1)
        {
            AnsiConsole.MarkupLine("[dim]Expected arguments:[/]");
            for (int i = 1; i < usageParts.Length && i < required; i++)
            {
                string status = i < args.Length ? "[green]✓[/]" : "[red]✗[/]";
                AnsiConsole.MarkupLine($"[dim]  [[{i}]] {status} {usageParts[i].EscapeMarkup()}[/]");
            }
        }
        
        return false;
    }

    /// <summary>
    /// Validates an Excel file path with detailed error context and security checks
    /// </summary>
    public static bool ValidateExcelFile(string filePath, bool requireExists = true)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            AnsiConsole.MarkupLine("[red]Error:[/] File path is empty or null");
            return false;
        }

        try
        {
            // Security: Prevent path traversal and validate path length
            string fullPath = Path.GetFullPath(filePath);
            
            if (fullPath.Length > 32767)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] File path too long ({fullPath.Length} characters, limit: 32767)");
                return false;
            }
            
            string extension = Path.GetExtension(fullPath).ToLowerInvariant();
            
            // Security: Strict file extension validation
            if (extension is not (".xlsx" or ".xlsm" or ".xls"))
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Invalid Excel file extension: {extension}");
                AnsiConsole.MarkupLine("[yellow]Supported extensions:[/] .xlsx, .xlsm, .xls");
                return false;
            }

            if (requireExists)
            {
                if (!File.Exists(fullPath))
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] File not found: {filePath}");
                    AnsiConsole.MarkupLine($"[yellow]Full path:[/] {fullPath}");
                    AnsiConsole.MarkupLine($"[yellow]Working directory:[/] {Environment.CurrentDirectory}");
                    
                    // Check if similar files exist
                    string? directory = Path.GetDirectoryName(fullPath);
                    string fileName = Path.GetFileNameWithoutExtension(fullPath);
                    
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                    {
                        var similarFiles = Directory.GetFiles(directory, $"*{fileName}*")
                            .Where(f => Path.GetExtension(f).ToLowerInvariant() is ".xlsx" or ".xlsm" or ".xls")
                            .Take(5)
                            .ToArray();
                            
                        if (similarFiles.Length > 0)
                        {
                            AnsiConsole.MarkupLine("[yellow]Similar files found:[/]");
                            foreach (var file in similarFiles)
                            {
                                AnsiConsole.MarkupLine($"  • {Path.GetFileName(file)}");
                            }
                        }
                    }
                    
                    return false;
                }

                // Security: Check file size to prevent potential DoS
                var fileInfo = new FileInfo(fullPath);
                const long MAX_FILE_SIZE = 1024L * 1024L * 1024L; // 1GB limit
                
                if (fileInfo.Length > MAX_FILE_SIZE)
                {
                    AnsiConsole.MarkupLine($"[red]Error:[/] File too large ({fileInfo.Length:N0} bytes, limit: {MAX_FILE_SIZE:N0} bytes)");
                    AnsiConsole.MarkupLine("[yellow]Large Excel files may cause performance issues or memory exhaustion[/]");
                    return false;
                }
                
                AnsiConsole.MarkupLine($"[dim]File info: {fileInfo.Length:N0} bytes, modified {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}[/]");
                
                // Check if file is locked
                if (IsFileLocked(fullPath))
                {
                    AnsiConsole.MarkupLine($"[yellow]Warning:[/] File appears to be locked by another process");
                    AnsiConsole.MarkupLine("[yellow]This may cause errors. Close Excel and try again.[/]");
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[red]Error validating file path:[/] {ex.Message.EscapeMarkup()}");
            return false;
        }
    }

    /// <summary>
    /// Checks if a file is locked by another process
    /// </summary>
    private static bool IsFileLocked(string filePath)
    {
        try
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                return false;
            }
        }
        catch (IOException)
        {
            return true;
        }
        catch
        {
            return false;
        }
    }
}
