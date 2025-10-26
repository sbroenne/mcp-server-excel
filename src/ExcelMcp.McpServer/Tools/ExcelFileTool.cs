using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core;
using Sbroenne.ExcelMcp.Core.Commands;

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
    [Description("Manage Excel files. Supports: create-empty, close-workbook. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelFile(
        [Description("Action to perform: create-empty, close-workbook")]
        string action,

        [Description("Excel file path (.xlsx or .xlsm extension)")]
        string excelPath,
        
        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var fileCommands = new FileCommands();

            switch (action.ToLowerInvariant())
            {
                case "create-empty":
                    // Determine if macro-enabled based on file extension
                    bool macroEnabled = excelPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase);
                    return await CreateEmptyFileAsync(fileCommands, excelPath, macroEnabled, batchId);

                case "close-workbook":
                    return CloseWorkbook(excelPath);

                default:
                    throw new ModelContextProtocol.McpException($"Unknown action '{action}'. Supported: create-empty, close-workbook");
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
    /// Note: File creation doesn't use batch sessions since it creates a new file.
    /// </summary>
    private static async Task<string> CreateEmptyFileAsync(FileCommands fileCommands, string excelPath, bool macroEnabled, string? batchId)
    {
        var extension = macroEnabled ? ".xlsm" : ".xlsx";
        if (!excelPath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            excelPath = Path.ChangeExtension(excelPath, extension);
        }

        // Note: CreateEmpty doesn't use batch session - it creates a new file
        // batchId is ignored for this operation
        var result = await fileCommands.CreateEmptyAsync(excelPath, overwriteIfExists: false);
            
        if (result.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                filePath = result.FilePath,
                macroEnabled,
                message = "Excel file created successfully",
                suggestedNextActions = new[]
                {
                    batchId != null 
                        ? $"Continue using batchId '{batchId}' for subsequent operations on this file"
                        : "Use begin_excel_batch to start a batch session for multiple operations",
                    "Use worksheet 'create' to add new worksheets",
                    "Use PowerQuery 'import' to add data transformations",
                    macroEnabled ? "Use VBA 'import' to add macro code" : "Use worksheet 'write' to populate data"
                },
                workflowHint = macroEnabled
                    ? "Macro-enabled file created. Next, add worksheets, Power Query, or VBA code."
                    : "Excel file created. Next, add worksheets and populate data."
            }, ExcelToolsBase.JsonOptions);
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                error = result.ErrorMessage,
                filePath = result.FilePath,
                suggestedNextActions = new[]
                {
                    "Check that the target directory exists and is writable",
                    "Verify the file doesn't already exist",
                    "Try a different file path"
                },
                workflowHint = "File creation failed. Ensure the path is valid and writable."
            }, ExcelToolsBase.JsonOptions);
        }
    }

    /// <summary>
    /// Closes the workbook (no-op with new single-instance architecture).
    /// LLM Pattern: This action is kept for backward compatibility but does nothing.
    /// With single-instance sessions, workbooks are automatically closed after each operation.
    /// </summary>
    private static string CloseWorkbook(string excelPath)
    {
        return JsonSerializer.Serialize(new
        {
            success = true,
            filePath = excelPath,
            message = "Workbook closure is automatic with single-instance architecture",
            suggestedNextActions = new[]
            {
                "Use 'excel_file' with action 'create-empty' to create new files",
                "Use other excel_* tools to work with files",
                "Each operation automatically manages its own Excel instance"
            },
            workflowHint = "With single-instance architecture, workbooks are automatically closed after each operation."
        }, ExcelToolsBase.JsonOptions);
    }
}
