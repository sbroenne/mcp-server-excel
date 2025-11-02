using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel worksheet lifecycle and appearance management tool for MCP server.
/// Handles worksheet creation, renaming, copying, deletion, tab colors, and visibility.
///
/// Data operations (read, write, clear) have been moved to ExcelRangeTool for unified range API.
///
/// LLM Usage Patterns:
/// - Use "list" to see all worksheets in a workbook
/// - Use "create" to add new worksheets
/// - Use "rename" to change worksheet names
/// - Use "copy" to duplicate worksheets
/// - Use "delete" to remove worksheets
/// - Use "set-tab-color" to color-code sheets (RGB 0-255 each)
/// - Use "get-tab-color" to read tab colors
/// - Use "clear-tab-color" to remove colors
/// - Use "set-visibility" to control sheet visibility (visible/hidden/veryhidden)
/// - Use "get-visibility" to check visibility state
/// - Use "show", "hide", "very-hide" as convenience methods
/// - Use excel_range tool for data operations (get-values, set-values, clear-*)
/// </summary>
[McpServerToolType]
public static class ExcelWorksheetTool
{
    /// <summary>
    /// Manage Excel worksheet lifecycle and appearance
    /// </summary>
    [McpServerTool(Name = "excel_worksheet")]
    [Description("Manage Excel worksheets: lifecycle, tab colors, visibility. Actions available as dropdown. Optional batchId for batch sessions.")]
    public static async Task<string> ExcelWorksheet(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        WorksheetAction action,

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
        [Description("New sheet name (for rename) or target sheet name (for copy)")]
        string? targetName = null,

        [Range(0, 255)]
        [Description("Red component (0-255) for set-tab-color action")]
        int? red = null,

        [Range(0, 255)]
        [Description("Green component (0-255) for set-tab-color action")]
        int? green = null,

        [Range(0, 255)]
        [Description("Blue component (0-255) for set-tab-color action")]
        int? blue = null,

        [RegularExpression("^(visible|hidden|veryhidden)$")]
        [Description("Visibility level for set-visibility action: visible (normal), hidden (user can unhide), veryhidden (requires code to unhide)")]
        string? visibility = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var sheetCommands = new SheetCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                WorksheetAction.List => await ListWorksheetsAsync(sheetCommands, excelPath, batchId),
                WorksheetAction.Create => await CreateWorksheetAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.Rename => await RenameWorksheetAsync(sheetCommands, excelPath, sheetName, targetName, batchId),
                WorksheetAction.Copy => await CopyWorksheetAsync(sheetCommands, excelPath, sheetName, targetName, batchId),
                WorksheetAction.Delete => await DeleteWorksheetAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.SetTabColor => await SetTabColorAsync(sheetCommands, excelPath, sheetName, red, green, blue, batchId),
                WorksheetAction.GetTabColor => await GetTabColorAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.ClearTabColor => await ClearTabColorAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.SetVisibility => await SetVisibilityAsync(sheetCommands, excelPath, sheetName, visibility, batchId),
                WorksheetAction.GetVisibility => await GetVisibilityAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.Show => await ShowAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.Hide => await HideAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.VeryHide => await VeryHideAsync(sheetCommands, excelPath, sheetName, batchId),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw;
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

            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

        // Add workflow guidance
        var sheetCount = result.Items?.Count ?? 0;
        return JsonSerializer.Serialize(new
        {
            success = result.Success,
            items = result.Items,
            suggestedNextActions = sheetCount == 0
                ? new[] { "Use worksheet 'create' to add your first worksheet" }
                : new[]
                {
                    "Use worksheet 'create' to add more worksheets",
                    "Use excel_range for data operations",
                    "Use worksheet 'rename' or 'set-tab-color' to organize sheets"
                },
            workflowHint = sheetCount == 0
                ? "Empty workbook. Create worksheets to store data."
                : $"Found {sheetCount} worksheet(s). Ready for data operations."
        }, ExcelToolsBase.JsonOptions);
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

        // Use workflow guidance with batch mode awareness
        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {

            throw new ModelContextProtocol.McpException($"create failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(new
        {
            success = result.Success,
            name = result.Name,
            suggestedNextActions = new[]
            {
                usedBatchMode
                    ? "Continue with more operations in this batch"
                    : "Use begin_excel_batch for multiple operations",
                $"Use excel_range to populate '{sheetName}' with data",
                "Use worksheet 'set-tab-color' to organize sheets visually"
            },
            workflowHint = $"Worksheet '{sheetName}' created. Ready for data operations."
        }, ExcelToolsBase.JsonOptions);
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

            throw new ModelContextProtocol.McpException($"rename failed for '{filePath}': {result.ErrorMessage}");
        }

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

            throw new ModelContextProtocol.McpException($"copy failed for '{filePath}': {result.ErrorMessage}");
        }

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

            throw new ModelContextProtocol.McpException($"delete failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === TAB COLOR OPERATIONS ===

    private static async Task<string> SetTabColorAsync(SheetCommands commands, string filePath, string? sheetName, int? red, int? green, int? blue, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for set-tab-color action");

        if (!red.HasValue || !green.HasValue || !blue.HasValue)
            throw new ModelContextProtocol.McpException("red, green, and blue values (0-255) are required for set-tab-color action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetTabColorAsync(batch, sheetName, red.Value, green.Value, blue.Value));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-tab-color failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetTabColorAsync(SheetCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for get-tab-color action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetTabColorAsync(batch, sheetName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-tab-color failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearTabColorAsync(SheetCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for clear-tab-color action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.ClearTabColorAsync(batch, sheetName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"clear-tab-color failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // === VISIBILITY OPERATIONS ===

    private static async Task<string> SetVisibilityAsync(SheetCommands commands, string filePath, string? sheetName, string? visibility, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for set-visibility action");

        if (string.IsNullOrEmpty(visibility))
            throw new ModelContextProtocol.McpException("visibility (visible|hidden|veryhidden) is required for set-visibility action");

        SheetVisibility visibilityLevel = visibility.ToLowerInvariant() switch
        {
            "visible" => SheetVisibility.Visible,
            "hidden" => SheetVisibility.Hidden,
            "veryhidden" => SheetVisibility.VeryHidden,
            _ => throw new ModelContextProtocol.McpException($"Invalid visibility value '{visibility}'. Use: visible, hidden, or veryhidden")
        };

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.SetVisibilityAsync(batch, sheetName, visibilityLevel));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-visibility failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetVisibilityAsync(SheetCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for get-visibility action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetVisibilityAsync(batch, sheetName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get-visibility failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ShowAsync(SheetCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for show action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.ShowAsync(batch, sheetName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"show failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> HideAsync(SheetCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for hide action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.HideAsync(batch, sheetName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"hide failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> VeryHideAsync(SheetCommands commands, string filePath, string? sheetName, string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for very-hide action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.VeryHideAsync(batch, sheetName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"very-hide failed: {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
