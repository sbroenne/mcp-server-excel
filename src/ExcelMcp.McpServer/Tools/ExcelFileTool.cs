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
    [McpServerTool(Name = "file")]
    [Description("Manage Excel files. Supports: create-empty, close-workbook.")]
    public static string File(
        [Description("Action to perform: create-empty, close-workbook")]
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
                message = "Excel file created successfully",
                suggestedNextActions = new[]
                {
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
    /// Closes the workbook in the pool, freeing up the instance slot.
    /// LLM Pattern: Use this when you're done working with a file to free up pool capacity.
    /// </summary>
    private static string CloseWorkbook(string excelPath)
    {
        // Close workbook in pool (if pooling is enabled)
        var pool = ExcelToolsPoolManager.Pool;
        if (pool != null)
        {
            pool.CloseWorkbook(excelPath);

            return JsonSerializer.Serialize(new
            {
                success = true,
                filePath = excelPath,
                message = "Workbook closed in pool. Instance slot freed for reuse.",
                suggestedNextActions = new[]
                {
                    "Pool capacity restored - you can now open other files",
                    "Use 'file' with action 'create-empty' to create new files",
                    "Use other excel_* tools to work with different files"
                },
                workflowHint = "Workbook closed. Pool instance slot is now available for other files."
            }, ExcelToolsBase.JsonOptions);
        }
        else
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                filePath = excelPath,
                message = "No pooling enabled - workbook close not needed",
                workflowHint = "Pooling is not enabled in this context. Workbook closure is automatic."
            }, ExcelToolsBase.JsonOptions);
        }
    }
}
