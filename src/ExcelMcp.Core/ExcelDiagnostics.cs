using System.Runtime.InteropServices;
using System.Text;
using Spectre.Console;

namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Enhanced Excel diagnostics and error reporting for coding agents
/// Provides comprehensive context when Excel operations fail
/// </summary>
public static class ExcelDiagnostics
{
    /// <summary>
    /// Captures comprehensive Excel environment and error context
    /// </summary>
    public static void ReportExcelError(Exception ex, string operation, string? filePath = null, dynamic? workbook = null, dynamic? excel = null)
    {
        var errorReport = new StringBuilder();
        errorReport.AppendLine($"Excel Operation Failed: {operation}");
        errorReport.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        errorReport.AppendLine();

        // Basic error information
        errorReport.AppendLine("=== ERROR DETAILS ===");
        errorReport.AppendLine($"Type: {ex.GetType().Name}");
        errorReport.AppendLine($"Message: {ex.Message}");
        errorReport.AppendLine($"HResult: 0x{ex.HResult:X8}");
        
        if (ex is COMException comEx)
        {
            errorReport.AppendLine($"COM Error Code: 0x{comEx.ErrorCode:X8}");
            errorReport.AppendLine($"COM Error Description: {GetComErrorDescription(comEx.ErrorCode)}");
        }

        if (ex.InnerException != null)
        {
            errorReport.AppendLine($"Inner Exception: {ex.InnerException.GetType().Name}");
            errorReport.AppendLine($"Inner Message: {ex.InnerException.Message}");
        }

        errorReport.AppendLine();

        // File context
        if (!string.IsNullOrEmpty(filePath))
        {
            errorReport.AppendLine("=== FILE CONTEXT ===");
            errorReport.AppendLine($"File Path: {filePath}");
            errorReport.AppendLine($"File Exists: {File.Exists(filePath)}");
            
            if (File.Exists(filePath))
            {
                var fileInfo = new FileInfo(filePath);
                errorReport.AppendLine($"File Size: {fileInfo.Length:N0} bytes");
                errorReport.AppendLine($"Last Modified: {fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");
                errorReport.AppendLine($"File Extension: {fileInfo.Extension}");
                errorReport.AppendLine($"Read Only: {fileInfo.IsReadOnly}");
                
                // Check if file is locked
                bool isLocked = IsFileLocked(filePath);
                errorReport.AppendLine($"File Locked: {isLocked}");
                
                if (isLocked)
                {
                    errorReport.AppendLine("WARNING: File appears to be locked by another process");
                    errorReport.AppendLine("SOLUTION: Close Excel and any other applications using this file");
                }
            }
            errorReport.AppendLine();
        }

        // Excel application context
        if (excel != null)
        {
            errorReport.AppendLine("=== EXCEL APPLICATION CONTEXT ===");
            try
            {
                errorReport.AppendLine($"Excel Version: {excel.Version ?? "Unknown"}");
                errorReport.AppendLine($"Excel Build: {excel.Build ?? "Unknown"}");
                errorReport.AppendLine($"Display Alerts: {excel.DisplayAlerts}");
                errorReport.AppendLine($"Visible: {excel.Visible}");
                errorReport.AppendLine($"Interactive: {excel.Interactive}");
                errorReport.AppendLine($"Calculation: {GetCalculationMode(excel.Calculation)}");
                
                dynamic workbooks = excel.Workbooks;
                errorReport.AppendLine($"Open Workbooks: {workbooks.Count}");
                
                // List open workbooks
                for (int i = 1; i <= Math.Min(workbooks.Count, 10); i++)
                {
                    try
                    {
                        dynamic wb = workbooks.Item(i);
                        errorReport.AppendLine($"  [{i}] {wb.Name} (Saved: {wb.Saved})");
                    }
                    catch
                    {
                        errorReport.AppendLine($"  [{i}] <Error accessing workbook>");
                    }
                }
                
                if (workbooks.Count > 10)
                {
                    errorReport.AppendLine($"  ... and {workbooks.Count - 10} more workbooks");
                }
            }
            catch (Exception diagEx)
            {
                errorReport.AppendLine($"Error gathering Excel context: {diagEx.Message}");
            }
            errorReport.AppendLine();
        }

        // Workbook context
        if (workbook != null)
        {
            errorReport.AppendLine("=== WORKBOOK CONTEXT ===");
            try
            {
                errorReport.AppendLine($"Workbook Name: {workbook.Name}");
                errorReport.AppendLine($"Full Name: {workbook.FullName}");
                errorReport.AppendLine($"Saved: {workbook.Saved}");
                errorReport.AppendLine($"Read Only: {workbook.ReadOnly}");
                errorReport.AppendLine($"Protected: {workbook.ProtectStructure}");
                
                dynamic worksheets = workbook.Worksheets;
                errorReport.AppendLine($"Worksheets: {worksheets.Count}");
                
                // List first few worksheets
                for (int i = 1; i <= Math.Min(worksheets.Count, 5); i++)
                {
                    try
                    {
                        dynamic ws = worksheets.Item(i);
                        errorReport.AppendLine($"  [{i}] {ws.Name} (Visible: {ws.Visible == -1})");
                    }
                    catch
                    {
                        errorReport.AppendLine($"  [{i}] <Error accessing worksheet>");
                    }
                }
                
                // Power Queries
                try
                {
                    dynamic queries = workbook.Queries;
                    errorReport.AppendLine($"Power Queries: {queries.Count}");
                    
                    for (int i = 1; i <= Math.Min(queries.Count, 5); i++)
                    {
                        try
                        {
                            dynamic query = queries.Item(i);
                            errorReport.AppendLine($"  [{i}] {query.Name}");
                        }
                        catch
                        {
                            errorReport.AppendLine($"  [{i}] <Error accessing query>");
                        }
                    }
                }
                catch
                {
                    errorReport.AppendLine("Power Queries: <Not accessible>");
                }
                
                // Named ranges
                try
                {
                    dynamic names = workbook.Names;
                    errorReport.AppendLine($"Named Ranges: {names.Count}");
                }
                catch
                {
                    errorReport.AppendLine("Named Ranges: <Not accessible>");
                }
            }
            catch (Exception diagEx)
            {
                errorReport.AppendLine($"Error gathering workbook context: {diagEx.Message}");
            }
            errorReport.AppendLine();
        }

        // System context
        errorReport.AppendLine("=== SYSTEM CONTEXT ===");
        errorReport.AppendLine($"OS: {Environment.OSVersion}");
        errorReport.AppendLine($"64-bit OS: {Environment.Is64BitOperatingSystem}");
        errorReport.AppendLine($"64-bit Process: {Environment.Is64BitProcess}");
        errorReport.AppendLine($"CLR Version: {Environment.Version}");
        errorReport.AppendLine($"Working Directory: {Environment.CurrentDirectory}");
        errorReport.AppendLine($"Available Memory: {GC.GetTotalMemory(false):N0} bytes");
        
        // Excel processes
        try
        {
            var excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            errorReport.AppendLine($"Excel Processes: {excelProcesses.Length}");
            
            foreach (var proc in excelProcesses.Take(5))
            {
                try
                {
                    errorReport.AppendLine($"  PID {proc.Id}: {proc.ProcessName} (Started: {proc.StartTime:HH:mm:ss})");
                }
                catch
                {
                    errorReport.AppendLine($"  PID {proc.Id}: <Process details unavailable>");
                }
            }
            
            if (excelProcesses.Length > 5)
            {
                errorReport.AppendLine($"  ... and {excelProcesses.Length - 5} more Excel processes");
            }
        }
        catch (Exception diagEx)
        {
            errorReport.AppendLine($"Error checking Excel processes: {diagEx.Message}");
        }

        errorReport.AppendLine();

        // Recommendations for coding agents
        errorReport.AppendLine("=== CODING AGENT RECOMMENDATIONS ===");
        
        if (ex is COMException comException)
        {
            var recommendations = GetComErrorRecommendations(comException.ErrorCode);
            foreach (var recommendation in recommendations)
            {
                errorReport.AppendLine($"• {recommendation}");
            }
        }
        else
        {
            errorReport.AppendLine("• Verify Excel is properly installed and accessible");
            errorReport.AppendLine("• Check file permissions and ensure file is not locked");
            errorReport.AppendLine("• Consider retrying the operation after a brief delay");
            errorReport.AppendLine("• Ensure all Excel applications are closed before retry");
        }

        errorReport.AppendLine();
        errorReport.AppendLine("=== STACK TRACE ===");
        errorReport.AppendLine(ex.StackTrace ?? "No stack trace available");

        // Output the comprehensive error report
        AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
        AnsiConsole.WriteLine();
        
        var panel = new Panel(errorReport.ToString().EscapeMarkup())
            .Header("[red bold]Detailed Excel Error Report for Coding Agent[/]")
            .BorderColor(Color.Red)
            .Padding(1, 1);
            
        AnsiConsole.Write(panel);
    }

    /// <summary>
    /// Gets human-readable description for COM error codes
    /// </summary>
    private static string GetComErrorDescription(int errorCode)
    {
        return unchecked((uint)errorCode) switch
        {
            0x800401E4 => "MK_E_SYNTAX - Moniker syntax error",
            0x80004005 => "E_FAIL - Unspecified failure",
            0x8007000E => "E_OUTOFMEMORY - Out of memory",
            0x80070005 => "E_ACCESSDENIED - Access denied",
            0x80070006 => "E_HANDLE - Invalid handle",
            0x8007000C => "E_UNEXPECTED - Unexpected failure",
            0x80004004 => "E_ABORT - Operation aborted",
            0x80004003 => "E_POINTER - Invalid pointer",
            0x80004002 => "E_NOINTERFACE - Interface not supported",
            0x80004001 => "E_NOTIMPL - Not implemented",
            0x8001010A => "RPC_E_SERVERCALL_RETRYLATER - Excel is busy, try again later",
            0x80010108 => "RPC_E_DISCONNECTED - Object disconnected from server",
            0x800706BE => "RPC_S_REMOTE_DISABLED - Remote procedure calls disabled",
            0x800706BA => "RPC_S_SERVER_UNAVAILABLE - RPC server unavailable",
            0x80131040 => "COR_E_FILENOTFOUND - File not found",
            0x80070002 => "ERROR_FILE_NOT_FOUND - System cannot find file",
            0x80070003 => "ERROR_PATH_NOT_FOUND - System cannot find path",
            0x80070020 => "ERROR_SHARING_VIOLATION - File is being used by another process",
            0x80030005 => "STG_E_ACCESSDENIED - Storage access denied",
            0x80030008 => "STG_E_INSUFFICIENTMEMORY - Insufficient memory",
            0x8003001D => "STG_E_WRITEFAULT - Disk write error",
            0x80030103 => "STG_E_CANTSAVE - Cannot save file",
            _ => $"Unknown COM error (0x{errorCode:X8})"
        };
    }

    /// <summary>
    /// Gets specific recommendations for COM error codes
    /// </summary>
    private static List<string> GetComErrorRecommendations(int errorCode)
    {
        var recommendations = new List<string>();
        
        switch (unchecked((uint)errorCode))
        {
            case 0x8001010A: // RPC_E_SERVERCALL_RETRYLATER
                recommendations.Add("Excel is busy - close any open dialogs in Excel");
                recommendations.Add("Wait 2-3 seconds and retry the operation");
                recommendations.Add("Ensure no other processes are accessing Excel");
                break;
                
            case 0x80070020: // ERROR_SHARING_VIOLATION
                recommendations.Add("File is locked by another process - close Excel and any file viewers");
                recommendations.Add("Check if file is open in another Excel instance");
                recommendations.Add("Use Task Manager to end all EXCEL.exe processes if needed");
                break;
                
            case 0x80070005: // E_ACCESSDENIED
                recommendations.Add("Run as Administrator if file is in protected location");
                recommendations.Add("Check file permissions and ensure write access");
                recommendations.Add("Verify file is not marked as read-only");
                break;
                
            case 0x80030103: // STG_E_CANTSAVE
                recommendations.Add("Check disk space availability");
                recommendations.Add("Verify target directory exists and is writable");
                recommendations.Add("Try saving to a different location");
                break;
                
            case 0x80004005: // E_FAIL
                recommendations.Add("Generic failure - check Excel installation");
                recommendations.Add("Try repairing Office installation");
                recommendations.Add("Restart Excel application");
                break;
                
            default:
                recommendations.Add("Check Excel installation and COM registration");
                recommendations.Add("Ensure Excel is not in compatibility mode");
                recommendations.Add("Verify file format matches extension (.xlsx/.xlsm)");
                break;
        }
        
        return recommendations;
    }

    /// <summary>
    /// Gets human-readable calculation mode
    /// </summary>
    private static string GetCalculationMode(dynamic calculation)
    {
        try
        {
            int mode = calculation;
            return mode switch
            {
                -4105 => "Automatic",
                -4135 => "Manual", 
                2 => "Automatic Except Tables",
                _ => $"Unknown ({mode})"
            };
        }
        catch
        {
            return "Unknown";
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

    /// <summary>
    /// Reports operation context for debugging
    /// </summary>
    public static void ReportOperationContext(string operation, string? filePath = null, params (string key, object? value)[] contextData)
    {
        var context = new StringBuilder();
        context.AppendLine($"Operation: {operation}");
        context.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
        
        if (!string.IsNullOrEmpty(filePath))
        {
            context.AppendLine($"File: {filePath}");
        }
        
        foreach (var (key, value) in contextData)
        {
            context.AppendLine($"{key}: {value ?? "null"}");
        }
        
        AnsiConsole.MarkupLine($"[dim]Debug Context:[/]");
        AnsiConsole.MarkupLine($"[dim]{context.ToString().EscapeMarkup()}[/]");
    }
}