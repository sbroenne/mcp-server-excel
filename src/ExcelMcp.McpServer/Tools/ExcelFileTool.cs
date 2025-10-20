using Sbroenne.ExcelMcp.Core.Commands;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
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
[McpServerToolType]
public static class ExcelFileTool
{
    /// <summary>
    /// Create new Excel files for automation workflows
    /// </summary>
    [McpServerTool(Name = "excel_file")]
    [Description("Manage Excel files. Supports: create-empty.")]
    public static string ExcelFile(
        [Description("Action to perform: create-empty")] 
        string action,
        
        [Description("Excel file path (.xlsx or .xlsm extension)")] 
        string excelPath)
    {
        try
        {
            var fileCommands = new FileCommands();

            switch (action.ToLowerInvariant())
            {
                case "create-empty":
                    // Determine if macro-enabled based on file extension
                    bool macroEnabled = excelPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);
                    return CreateEmptyFile(fileCommands, excelPath, macroEnabled);
                    
                default:
                    throw new ModelContextProtocol.McpException($"Unknown action '{action}'. Supported: create-empty");
            }
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action, excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    /// <summary>
    /// Creates a new empty Excel file (.xlsx or .xlsm based on macroEnabled flag).
    /// LLM Pattern: Use this when you need a fresh Excel workbook for automation.
    /// </summary>
    private static string CreateEmptyFile(FileCommands fileCommands, string excelPath, bool macroEnabled)
    {
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        if (!excelPath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            excelPath = Path.ChangeExtension(excelPath, extension);
        }

        var result = fileCommands.CreateEmpty(excelPath, overwriteIfExists: false);
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