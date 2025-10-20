using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel file management tool for MCP server.
/// Handles Excel file creation for automation workflows.
/// 
/// LLM Usage Pattern:
/// - Use "create-empty" for new Excel files in automation workflows
/// - File validation and existence checks can be done with standard file system operations
/// </summary>
public static class ExcelFileTool
{
    /// <summary>
    /// Create new Excel files for automation workflows
    /// </summary>
    [McpServerTool(Name = "excel_file")]
        [Description("Manage Excel files. Supports: create-empty.")]
    public static string ExcelFile(
        [Description("Action to perform: create-empty")] string action,
        [Description("Excel file path (.xlsx or .xlsm extension)")] string filePath,
        [Description("Optional: macro-enabled flag for create-empty (default: false)")] bool macroEnabled = false)
    {
        try
        {
            var fileCommands = new FileCommands();

            return action.ToLowerInvariant() switch
            {
                "create-empty" => CreateEmptyFile(fileCommands, filePath, macroEnabled),
                _ => ExcelToolsBase.CreateUnknownActionError(action, "create-empty")
            };
        }
        catch (Exception ex)
        {
            return ExcelToolsBase.CreateExceptionError(ex, action, filePath);
        }
    }

    /// <summary>
    /// Creates a new empty Excel file (.xlsx or .xlsm based on macroEnabled flag).
    /// LLM Pattern: Use this when you need a fresh Excel workbook for automation.
    /// </summary>
    private static string CreateEmptyFile(FileCommands fileCommands, string filePath, bool macroEnabled)
    {
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        if (!filePath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            filePath = Path.ChangeExtension(filePath, extension);
        }

        var result = fileCommands.CreateEmpty(filePath, overwriteIfExists: false);
        if (result.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                filePath = result.FilePath,
                macroEnabled,
                message = "Excel file created successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                error = result.ErrorMessage,
                filePath = result.FilePath
            }, ExcelToolsBase.JsonOptions);
        }
    }
}