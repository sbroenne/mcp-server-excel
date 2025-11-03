using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.McpServer.Models;

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
    [Description("Manage Excel files. Actions available as dropdown: CreateEmpty, CloseWorkbook, Test. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelFile(
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        FileAction action,

        [Description("Excel file path (.xlsx or .xlsm extension)")]
        string excelPath,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var fileCommands = new FileCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                FileAction.CreateEmpty => await CreateEmptyFileAsync(fileCommands, excelPath,
                    excelPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase), batchId),
                FileAction.CloseWorkbook => CloseWorkbook(excelPath),
                FileAction.Test => await TestFileAsync(fileCommands, excelPath),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
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
            throw new ModelContextProtocol.McpException($"create-empty failed for '{excelPath}': {result.ErrorMessage}");
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

    /// <summary>
    /// Tests if an Excel file exists and is valid.
    /// LLM Pattern: Use this for discovery/connectivity testing and validation before operations.
    /// This is a lightweight check that doesn't open the file with Excel COM.
    /// </summary>
    private static async Task<string> TestFileAsync(FileCommands fileCommands, string excelPath)
    {
        var result = await fileCommands.TestFileAsync(excelPath);

        if (result.Success)
        {
            return JsonSerializer.Serialize(new
            {
                success = true,
                filePath = result.FilePath,
                exists = result.Exists,
                isValid = result.IsValid,
                extension = result.Extension,
                size = result.Size,
                lastModified = result.LastModified,
                message = "File exists and is a valid Excel file",
                suggestedNextActions = new[]
                {
                    "Use excel_worksheet to manage worksheets",
                    "Use excel_powerquery to manage Power Query connections",
                    "Use excel_vba to manage VBA macros",
                    "Use begin_excel_batch for multi-operation workflows"
                },
                workflowHint = "File is ready for Excel operations."
            }, ExcelToolsBase.JsonOptions);
        }
        else
        {
            throw new ModelContextProtocol.McpException($"test failed for '{excelPath}': {result.ErrorMessage}");
        }
    }
}
