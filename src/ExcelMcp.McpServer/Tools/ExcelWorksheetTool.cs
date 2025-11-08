using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call

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
[SuppressMessage("Performance", "CA1861:Avoid constant arrays as arguments", Justification = "Simple workflow arrays in sealed static class")]
public static class ExcelWorksheetTool
{
    /// <summary>
    /// Manage Excel worksheet lifecycle and appearance
    /// </summary>
    [McpServerTool(Name = "excel_worksheet")]
    [Description(@"Manage Excel worksheets: lifecycle, tab colors, visibility.

TAB COLORS (set-tab-color):
- RGB values: 0-255 for red, green, blue components
- Example: red=255, green=0, blue=0 for red tab

VISIBILITY LEVELS (set-visibility):
- 'visible': Normal sheet (default)
- 'hidden': User can unhide via Excel UI (right-click → Unhide)
- 'veryhidden': Requires code to unhide (cannot unhide via UI)

RELATED TOOLS:
- excel_range: For ALL data operations (get-values, set-values, formulas, clear, etc.)
- excel_table: For structured data with AutoFilter

Optional batchId for batch sessions.")]
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

            // Expression switch pattern for audit compliance
            return action switch
            {
                WorksheetAction.List => await ListAsync(sheetCommands, excelPath, batchId),
                WorksheetAction.Create => await CreateAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.Rename => await RenameAsync(sheetCommands, excelPath, sheetName, targetName, batchId),
                WorksheetAction.Copy => await CopyAsync(sheetCommands, excelPath, sheetName, targetName, batchId),
                WorksheetAction.Delete => await DeleteAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.SetTabColor => await SetTabColorAsync(sheetCommands, excelPath, sheetName, red, green, blue, batchId),
                WorksheetAction.GetTabColor => await GetTabColorAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.ClearTabColor => await ClearTabColorAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.SetVisibility => await SetVisibilityAsync(sheetCommands, excelPath, sheetName, visibility, batchId),
                WorksheetAction.GetVisibility => await GetVisibilityAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.Show => await ShowAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.Hide => await HideAsync(sheetCommands, excelPath, sheetName, batchId),
                WorksheetAction.VeryHide => await VeryHideAsync(sheetCommands, excelPath, sheetName, batchId),
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
            throw;
        }
    }

    // === PRIVATE HELPER METHODS ===

    private static async Task<string> ListAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            sheetCommands.ListAsync);

        var count = result.Worksheets?.Count ?? 0;
        var inBatch = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            success = result.Success,
            worksheets = result.Worksheets,
            workflowHint = $"Found {count} worksheet(s). Use excel_range for data operations.",
            suggestedNextActions = count == 0
                ? new[] { "Workbook is empty - this shouldn't happen. Check file integrity." }
                :
                [
                "Use excel_range for data operations (get-values, set-values, clear-*)",
                "Use 'create' to add new worksheets",
                "Use 'set-tab-color' to organize sheets visually",
                inBatch ? "Continue batch operations" : count > 3 ? "Use excel_batch for multiple sheet operations (faster)" : "Use 'rename' or 'copy' to manage sheets"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for create action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.CreateAsync(batch, sheetName));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Worksheet '{sheetName}' created successfully.",
            suggestedNextActions = new[]
            {
                "Use excel_range 'set-values' to add data to the new sheet",
                "Use 'set-tab-color' to color-code this sheet",
                usedBatchMode ? "Create more worksheets in this batch" : "Creating multiple sheets? Use excel_batch (faster)"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RenameAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? targetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ModelContextProtocol.McpException("sheetName and targetName are required for rename action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.RenameAsync(batch, sheetName, targetName));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Worksheet '{sheetName}' renamed to '{targetName}' successfully."
                : $"Failed to rename worksheet: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Update any references to the old sheet name in formulas or code",
                    "Use excel_range to access the renamed sheet's data",
                    usedBatchMode ? "Continue renaming other sheets in this batch" : "Renaming multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the original sheet name exists using 'list' action",
                    "Check that the target name doesn't conflict with an existing sheet",
                    "Ensure the target name follows Excel naming rules (no special characters like [ ] : \\ / * ?)"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CopyAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? targetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ModelContextProtocol.McpException("sheetName and targetName are required for copy action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.CopyAsync(batch, sheetName, targetName));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Worksheet '{sheetName}' copied to '{targetName}' successfully."
                : $"Failed to copy worksheet: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Modify the copied sheet using excel_range (set-values, set-formulas)",
                    "Use 'set-tab-color' to visually distinguish the copy",
                    usedBatchMode ? "Copy more worksheets in this batch" : "Copying multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the source sheet name exists using 'list' action",
                    "Check that the target name doesn't conflict with an existing sheet",
                    "Ensure the target name follows Excel naming rules"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for delete action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.DeleteAsync(batch, sheetName));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Worksheet '{sheetName}' deleted successfully. Data is permanently removed."
                : $"Failed to delete worksheet: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Verify remaining worksheets using 'list' action",
                    "Check for broken references in formulas or VBA code",
                    usedBatchMode ? "Delete more worksheets in this batch" : "Deleting multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Check if workbook has only one sheet (Excel requires at least one)",
                    "Ensure the sheet is not protected"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetTabColorAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        int? red,
        int? green,
        int? blue,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for set-tab-color action");

        if (!red.HasValue)
            throw new ModelContextProtocol.McpException("red value (0-255) is required for set-tab-color action");
        if (!green.HasValue)
            throw new ModelContextProtocol.McpException("green value (0-255) is required for set-tab-color action");
        if (!blue.HasValue)
            throw new ModelContextProtocol.McpException("blue value (0-255) is required for set-tab-color action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.SetTabColorAsync(batch, sheetName, red.Value, green.Value, blue.Value));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);
        string hexColor = $"#{red.Value:X2}{green.Value:X2}{blue.Value:X2}";

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Tab color set to {hexColor} (RGB: {red.Value}, {green.Value}, {blue.Value}) for sheet '{sheetName}'."
                : $"Failed to set tab color: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'get-tab-color' to verify the color was applied",
                    "Apply consistent colors to related sheets for organization",
                    usedBatchMode ? "Set colors for more sheets in this batch" : "Coloring multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Check RGB values are in range 0-255",
                    "Use 'clear-tab-color' to remove color if needed"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetTabColorAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for get-tab-color action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await sheetCommands.GetTabColorAsync(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.HasColor,
            result.Red,
            result.Green,
            result.Blue,
            result.HexColor,
            result.ErrorMessage,
            workflowHint = result.Success
                ? (result.HasColor
                    ? $"Sheet '{sheetName}' has tab color: {result.HexColor} (RGB: {result.Red}, {result.Green}, {result.Blue})."
                    : $"Sheet '{sheetName}' has no tab color set (default).")
                : $"Failed to get tab color: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? (result.HasColor
                    ? new[]
                    {
                        "Use 'clear-tab-color' to remove the color",
                        "Use 'set-tab-color' to change the color",
                        "Check other sheets' colors for consistent organization"
                    }
                    :
                    [
                        "Use 'set-tab-color' to add a color for visual organization",
                        "Apply consistent colors to related sheets",
                        "Use colors to categorize sheets (e.g., red for important, blue for data)"
                    ])
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Check if the workbook is accessible",
                    "Retry the operation"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ClearTabColorAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for clear-tab-color action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.ClearTabColorAsync(batch, sheetName));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Tab color cleared for sheet '{sheetName}' (reset to default)."
                : $"Failed to clear tab color: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'get-tab-color' to verify the color was removed",
                    "Use 'set-tab-color' to apply a new color",
                    usedBatchMode ? "Clear colors from more sheets in this batch" : "Clearing multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Check if the sheet already has no color set",
                    "Retry the operation"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> SetVisibilityAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? visibility,
        string? batchId)
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
            excelPath,
            save: true,
            async (batch) => await sheetCommands.SetVisibilityAsync(batch, sheetName, visibilityLevel));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Visibility set to '{visibility}' for sheet '{sheetName}'."
                : $"Failed to set visibility: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'get-visibility' to verify the visibility level",
                    visibilityLevel == SheetVisibility.Hidden ? "Users can unhide this sheet via Excel UI" : (visibilityLevel == SheetVisibility.VeryHidden ? "Only code can unhide this sheet (good for protection)" : "Sheet is now visible in workbook"),
                    usedBatchMode ? "Set visibility for more sheets in this batch" : "Managing multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Ensure visibility value is: visible, hidden, or veryhidden",
                    "Check if workbook has at least one visible sheet"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetVisibilityAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for get-visibility action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await sheetCommands.GetVisibilityAsync(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Visibility,
            result.VisibilityName,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sheet '{sheetName}' visibility is '{result.VisibilityName}'."
                : $"Failed to get visibility: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? (result.Visibility == SheetVisibility.Visible
                    ? new[]
                    {
                        "Use 'hide' or 'very-hide' to hide this sheet",
                        "Sheet is currently visible in the workbook",
                        "Use 'set-visibility' for more control over visibility level"
                    }
                    : result.Visibility == SheetVisibility.Hidden
                        ?
                        [
                            "Use 'show' to make this sheet visible",
                            "Users can unhide this sheet via Excel UI",
                            "Use 'very-hide' for stronger protection"
                        ]
                        :
                        [
                            "Use 'show' to make this sheet visible",
                            "This sheet is very hidden - only code can unhide it",
                            "Good for protecting calculation or configuration sheets"
                        ])
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Check if the workbook is accessible",
                    "Retry the operation"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ShowAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for show action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.ShowAsync(batch, sheetName));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sheet '{sheetName}' is now visible in the workbook."
                : $"Failed to show sheet: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'get-visibility' to verify the sheet is visible",
                    "Access the sheet's data using excel_range",
                    usedBatchMode ? "Show more sheets in this batch" : "Showing multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Check if the sheet is already visible",
                    "Ensure the sheet is not protected"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> HideAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for hide action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.HideAsync(batch, sheetName));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sheet '{sheetName}' is now hidden (users can unhide via Excel UI)."
                : $"Failed to hide sheet: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'get-visibility' to verify the sheet is hidden",
                    "Users can unhide this sheet via Excel: Right-click sheet tab → Unhide",
                    "Use 'very-hide' for stronger protection (requires code to unhide)",
                    usedBatchMode ? "Hide more sheets in this batch" : "Hiding multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Check if workbook has at least one visible sheet",
                    "Ensure the sheet is not protected"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> VeryHideAsync(
        SheetCommands sheetCommands,
        string excelPath,
        string? sheetName,
        string? batchId)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for very-hide action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await sheetCommands.VeryHideAsync(batch, sheetName));

        bool usedBatchMode = !string.IsNullOrEmpty(batchId);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sheet '{sheetName}' is now very hidden (requires code to unhide)."
                : $"Failed to very hide sheet: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'get-visibility' to verify the sheet is very hidden",
                    "This sheet cannot be unhidden via Excel UI - only via code",
                    "Good for protecting calculation, configuration, or sensitive sheets",
                    usedBatchMode ? "Very hide more sheets in this batch" : "Protecting multiple sheets? Use excel_batch (faster)"
                }
                :
                [
                    "Verify the sheet name exists using 'list' action",
                    "Check if workbook has at least one visible sheet",
                    "Ensure the sheet is not protected"
                ]
        }, ExcelToolsBase.JsonOptions);
    }
}
