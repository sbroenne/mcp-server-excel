using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel worksheet lifecycle management tool for MCP server.
/// Handles worksheet creation, renaming, copying, and deletion.
/// 
/// Data operations (read, write, clear) have been moved to ExcelRangeTool for unified range API.
///
/// LLM Usage Patterns:
/// - Use "list" to see all worksheets in a workbook
/// - Use "create" to add new worksheets
/// - Use "rename" to change worksheet names
/// - Use "copy" to duplicate worksheets
/// - Use "delete" to remove worksheets
/// - Use excel_range tool for data operations (get-values, set-values, clear-*)
/// </summary>
[McpServerToolType]
public static class ExcelWorksheetTool
{
    /// <summary>
    /// Manage Excel worksheet lifecycle - create, rename, copy, delete sheets
    /// </summary>
    [McpServerTool(Name = "excel_worksheet")]
    [Description("Manage Excel worksheet lifecycle. Supports: list, create, rename, copy, delete. Use excel_range for data operations. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelWorksheet(
        [Required]
        [RegularExpression("^(list|create|rename|copy|delete)$")]
        [Description("Action: list, create, rename, copy, delete")]
        string action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Worksheet name (required for most actions)")]
        string? sheetName = null,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("New sheet name (for rename) or source sheet name (for copy)")]
        string? targetName = null,
        
        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var sheetCommands = new SheetCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => await ListWorksheetsAsync(sheetCommands, excelPath, batchId),
                "create" => await CreateWorksheetAsync(sheetCommands, excelPath, sheetName, batchId),
                "rename" => await RenameWorksheetAsync(sheetCommands, excelPath, sheetName, targetName, batchId),
                "copy" => await CopyWorksheetAsync(sheetCommands, excelPath, sheetName, targetName, batchId),
                "delete" => await DeleteWorksheetAsync(sheetCommands, excelPath, sheetName, batchId),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list, create, rename, copy, delete")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            throw new ModelContextProtocol.McpException($"Unexpected error in excel_worksheet action '{action}': {ex.Message}");
        }
    }

    private static async Task<string> ListWorksheetsAsync(SheetCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListAsync(batch));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the Excel file exists and is accessible",
                "Verify the file path is correct"
            };
            result.WorkflowHint = "List failed. Ensure the file exists and retry.";
            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use excel_range tool to read data from a worksheet",
            "Use 'create' to add a new worksheet",
            "Use 'delete' to remove a worksheet"
        };
        result.WorkflowHint = "Worksheets listed. Next, use excel_range for data or manage sheets.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateWorksheetAsync(SheetCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for create action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CreateAsync(batch, sheetName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet name doesn't already exist",
                "Verify the worksheet name is valid",
                "Use 'list' to see existing worksheets"
            };
            result.WorkflowHint = "Create failed. Ensure the worksheet name is unique and valid.";
            throw new ModelContextProtocol.McpException($"create failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use excel_range 'set-values' to populate the new worksheet",
            "Use 'list' to verify worksheet exists",
            "Use PowerQuery to load data into the sheet"
        };
        result.WorkflowHint = "Worksheet created successfully. Next, populate with data using excel_range.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RenameWorksheetAsync(SheetCommands commands, string filePath, string? sheetName, string? targetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ModelContextProtocol.McpException("sheetName and targetName are required for rename action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.RenameAsync(batch, sheetName, targetName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the source worksheet exists",
                "Verify the target name doesn't already exist",
                "Use 'list' to see available worksheets"
            };
            result.WorkflowHint = "Rename failed. Ensure the source exists and target is unique.";
            throw new ModelContextProtocol.McpException($"rename failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list' to verify the rename",
            "Use excel_range to access data in the renamed worksheet",
            "Update any formulas referencing the old name"
        };
        result.WorkflowHint = "Worksheet renamed successfully. Next, verify and update references.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CopyWorksheetAsync(SheetCommands commands, string filePath, string? sheetName, string? targetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ModelContextProtocol.McpException("sheetName and targetName are required for copy action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CopyAsync(batch, sheetName, targetName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the source worksheet exists",
                "Verify the target name doesn't already exist",
                "Use 'list' to see available worksheets"
            };
            result.WorkflowHint = "Copy failed. Ensure the source exists and target is unique.";
            throw new ModelContextProtocol.McpException($"copy failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list' to verify the copy",
            "Use excel_range to access data in the copied worksheet",
            "Modify the copied sheet independently using excel_range"
        };
        result.WorkflowHint = "Worksheet copied successfully. Next, verify and modify as needed.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteWorksheetAsync(SheetCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for delete action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteAsync(batch, sheetName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the worksheet exists",
                "Verify the worksheet is not the only sheet in the workbook",
                "Use 'list' to see available worksheets"
            };
            result.WorkflowHint = "Delete failed. Ensure the worksheet exists and is not the last sheet.";
            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list' to verify the deletion",
            "Update any formulas referencing the deleted sheet",
            "Review remaining worksheets"
        };
        result.WorkflowHint = "Worksheet deleted successfully. Next, verify and update references.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
