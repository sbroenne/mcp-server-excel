using Spectre.Console;
using static Sbroenne.ExcelMcp.Core.ExcelHelper;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// File management commands implementation
/// </summary>
public class FileCommands : IFileCommands
{
    /// <inheritdoc />
    public int CreateEmpty(string[] args)
    {
        if (!ValidateArgs(args, 2, "create-empty <file.xlsx|file.xlsm>")) return 1;

        string filePath = Path.GetFullPath(args[1]);
        
        // Validate file extension
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension != ".xlsx" && extension != ".xlsm")
        {
            AnsiConsole.MarkupLine("[red]Error:[/] File must have .xlsx or .xlsm extension");
            AnsiConsole.MarkupLine("[yellow]Tip:[/] Use .xlsm for macro-enabled workbooks");
            return 1;
        }
        
        // Check if file already exists
        if (File.Exists(filePath))
        {
            AnsiConsole.MarkupLine($"[yellow]Warning:[/] File already exists: {filePath}");
            
            // Ask for confirmation to overwrite
            if (!AnsiConsole.Confirm("Do you want to overwrite the existing file?"))
            {
                AnsiConsole.MarkupLine("[dim]Operation cancelled.[/]");
                return 1;
            }
        }

        // Ensure directory exists
        string? directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            try
            {
                Directory.CreateDirectory(directory);
                AnsiConsole.MarkupLine($"[dim]Created directory: {directory}[/]");
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error:[/] Failed to create directory: {ex.Message.EscapeMarkup()}");
                return 1;
            }
        }

        try
        {
            // Create Excel workbook with COM automation
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                AnsiConsole.MarkupLine("[red]Error:[/] Excel is not installed. Cannot create Excel files.");
                return 1;
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
                    AnsiConsole.MarkupLine($"[green]✓[/] Created macro-enabled Excel workbook: [cyan]{Path.GetFileName(filePath)}[/]");
                }
                else
                {
                    // Save as regular workbook (format 51)
                    workbook.SaveAs(filePath, 51);
                    AnsiConsole.MarkupLine($"[green]✓[/] Created Excel workbook: [cyan]{Path.GetFileName(filePath)}[/]");
                }
                
                workbook.Close(false);
                AnsiConsole.MarkupLine($"[dim]Full path: {filePath}[/]");
                
                return 0;
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
            AnsiConsole.MarkupLine($"[red]Error:[/] Failed to create Excel file: {ex.Message.EscapeMarkup()}");
            return 1;
        }
    }
}
