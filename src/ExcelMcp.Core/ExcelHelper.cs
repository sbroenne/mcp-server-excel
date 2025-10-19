using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Helper class for Excel COM automation with proper resource management
/// </summary>
public static class ExcelHelper
{
    /// <summary>
    /// Executes an action with Excel COM automation using proper resource management
    /// </summary>
    /// <typeparam name="T">Return type of the action</typeparam>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes to the file</param>
    /// <param name="action">Action to execute with Excel application and workbook</param>
    /// <returns>Result of the action</returns>
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
            catch
            {
                // Propagate exceptions with original context
                throw;
            }

            // Save if requested
            if (save && workbook != null)
            {
                try
                {
                    workbook.Save();
                }
                catch
                {
                    // Propagate save exceptions
                    throw;
                }
            }

            return result;
        }
        catch
        {
            // Propagate exceptions to caller
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

    /// <summary>
    /// Finds a Power Query by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="queryName">Name of the query to find</param>
    /// <returns>The query COM object if found, null otherwise</returns>
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

    /// <summary>
    /// Finds a named range by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="name">Name of the named range to find</param>
    /// <returns>The named range COM object if found, null otherwise</returns>
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

    /// <summary>
    /// Finds a worksheet by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="sheetName">Name of the worksheet to find</param>
    /// <returns>The worksheet COM object if found, null otherwise</returns>
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

}
