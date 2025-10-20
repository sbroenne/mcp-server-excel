using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel worksheet management tool for MCP server.
/// Handles worksheet operations, data reading/writing, and sheet management.
/// 
/// LLM Usage Patterns:
/// - Use "list" to see all worksheets in a workbook
/// - Use "read" to extract data from worksheet ranges
/// - Use "write" to populate worksheets from CSV files
/// - Use "create" to add new worksheets
/// - Use "rename" to change worksheet names  
/// - Use "copy" to duplicate worksheets
/// - Use "delete" to remove worksheets
/// - Use "clear" to empty worksheet ranges
/// - Use "append" to add data to existing worksheet content
/// </summary>
public static class ExcelWorksheetTool
{
    /// <summary>
    /// Manage Excel worksheets - data operations, sheet management, and content manipulation
    /// </summary>
    [McpServerTool(Name = "excel_worksheet")]  
    [Description("Manage Excel worksheets and data. Supports: list, read, write, create, rename, copy, delete, clear, append.")]
    public static string ExcelWorksheet(
        [Required]
        [RegularExpression("^(list|read|write|create|rename|copy|delete|clear|append)$")]
        [Description("Action: list, read, write, create, rename, copy, delete, clear, append")] 
        string action,
        
        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")] 
        string excelPath,
        
        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Worksheet name (required for most actions)")] 
        string? sheetName = null,
        
        [Description("Excel range (e.g., 'A1:D10' for read/clear) or CSV file path (for write/append)")] 
        string? range = null,
        
        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("New sheet name (for rename) or source sheet name (for copy)")] 
        string? targetName = null)
    {
        try
        {
            var sheetCommands = new SheetCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListWorksheets(sheetCommands, excelPath),
                "read" => ReadWorksheet(sheetCommands, excelPath, sheetName, range),
                "write" => WriteWorksheet(sheetCommands, excelPath, sheetName, range),
                "create" => CreateWorksheet(sheetCommands, excelPath, sheetName),
                "rename" => RenameWorksheet(sheetCommands, excelPath, sheetName, targetName),
                "copy" => CopyWorksheet(sheetCommands, excelPath, sheetName, targetName),
                "delete" => DeleteWorksheet(sheetCommands, excelPath, sheetName),
                "clear" => ClearWorksheet(sheetCommands, excelPath, sheetName, range),
                "append" => AppendWorksheet(sheetCommands, excelPath, sheetName, range),
                _ => ExcelToolsBase.CreateUnknownActionError(action, 
                    "list", "read", "write", "create", "rename", "copy", "delete", "clear", "append")
            };
        }
        catch (Exception ex)
        {
            return ExcelToolsBase.CreateExceptionError(ex, action, excelPath);
        }
    }

    private static string ListWorksheets(SheetCommands commands, string filePath)
    {
        var result = commands.List(filePath);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ReadWorksheet(SheetCommands commands, string filePath, string? sheetName, string? range)
    {
        if (string.IsNullOrEmpty(sheetName))
            return JsonSerializer.Serialize(new { error = "sheetName is required for read action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Read(filePath, sheetName, range ?? "");
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string WriteWorksheet(SheetCommands commands, string filePath, string? sheetName, string? dataPath)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(dataPath))
            return JsonSerializer.Serialize(new { error = "sheetName and range (CSV file path) are required for write action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Write(filePath, sheetName, dataPath);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string CreateWorksheet(SheetCommands commands, string filePath, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            return JsonSerializer.Serialize(new { error = "sheetName is required for create action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Create(filePath, sheetName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RenameWorksheet(SheetCommands commands, string filePath, string? sheetName, string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            return JsonSerializer.Serialize(new { error = "sheetName and targetName are required for rename action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Rename(filePath, sheetName, targetName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string CopyWorksheet(SheetCommands commands, string filePath, string? sheetName, string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            return JsonSerializer.Serialize(new { error = "sheetName and targetName are required for copy action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Copy(filePath, sheetName, targetName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteWorksheet(SheetCommands commands, string filePath, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            return JsonSerializer.Serialize(new { error = "sheetName is required for delete action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Delete(filePath, sheetName);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ClearWorksheet(SheetCommands commands, string filePath, string? sheetName, string? range)
    {
        if (string.IsNullOrEmpty(sheetName))
            return JsonSerializer.Serialize(new { error = "sheetName is required for clear action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Clear(filePath, sheetName, range ?? "");
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string AppendWorksheet(SheetCommands commands, string filePath, string? sheetName, string? dataPath)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(dataPath))
            return JsonSerializer.Serialize(new { error = "sheetName and range (CSV file path) are required for append action" }, ExcelToolsBase.JsonOptions);

        var result = commands.Append(filePath, sheetName, dataPath);
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}