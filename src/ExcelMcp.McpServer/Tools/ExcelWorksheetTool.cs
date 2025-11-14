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
/// MCP tool for Excel worksheet lifecycle and appearance (create, rename, copy, delete, tab colors, visibility).
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

REQUIRED WORKFLOW:
- Use excel_file(action: 'open') first to get a sessionId
- Pass sessionId to all worksheet actions
- Use excel_file(action: 'save') to persist changes
- Use excel_file(action: 'close') to end the session (does NOT save)

TAB COLORS (set-tab-color):
- RGB values: 0-255 for red, green, blue components
- Example: red=255, green=0, blue=0 for red tab
")]
    public static async Task<string> ExcelWorksheet(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        WorksheetAction action,

        [Required]
        [Description("Active Excel session ID from excel_file 'open' action")]
        string sessionId,

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
        string? visibility = null)
    {
        try
        {
            var sheetCommands = new SheetCommands();

            // Expression switch pattern for audit compliance
            return action switch
            {
                WorksheetAction.List => await ListAsync(sheetCommands, sessionId),
                WorksheetAction.Create => await CreateAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.Delete => await DeleteAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.Rename => await RenameAsync(sheetCommands, sessionId, sheetName, targetName),
                WorksheetAction.Copy => await CopyAsync(sheetCommands, sessionId, sheetName, targetName),
                WorksheetAction.SetTabColor => await SetTabColorAsync(sheetCommands, sessionId, sheetName, red, green, blue),
                WorksheetAction.GetTabColor => await GetTabColorAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.ClearTabColor => await ClearTabColorAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.SetVisibility => await SetVisibilityAsync(sheetCommands, sessionId, sheetName, visibility),
                WorksheetAction.GetVisibility => await GetVisibilityAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.Show => await ShowAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.Hide => await HideAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.VeryHide => await VeryHideAsync(sheetCommands, sessionId, sheetName),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (TimeoutException ex)
        {
            var result = new
            {
                success = false,
                errorMessage = ex.Message,
                operationContext = new Dictionary<string, object>
                {
                    { "OperationType", "excel_worksheet" },
                    { "Action", action.ToActionString() },
                    { "TimeoutReached", true }
                },
                isRetryable = !ex.Message.Contains("maximum timeout", StringComparison.OrdinalIgnoreCase),
                retryGuidance = ex.Message.Contains("maximum timeout", StringComparison.OrdinalIgnoreCase)
                    ? "Maximum timeout reached. Check workbook state manually."
                    : "Retry acceptable if issue is transient.",
                suggestedNextActions = new List<string>
                {
                    "Check if Excel is showing a dialog or prompt",
                    "Verify data source connectivity if operation touches external data",
                    "For large workbooks, operation may need more time"
                }
            };

            return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString());
            throw;
        }
    }

    // === PRIVATE HELPER METHODS ===

    private static async Task<string> ListAsync(
        SheetCommands sheetCommands,
        string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            sheetCommands.ListAsync);
        var count = result.Worksheets?.Count ?? 0;
        var inSession = !string.IsNullOrEmpty(sessionId);

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
                inSession ? "Continue working in this session" : "Use excel_file 'open' to start a session before worksheet operations"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for create action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.CreateAsync(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Worksheet '{sheetName}' created successfully.",
            suggestedNextActions = new[]
            {
                "Use excel_range 'set-values' to add data to the new sheet",
                "Use 'set-tab-color' to color-code this sheet",
                "Creating multiple sheets? Keep reusing this session for best performance"
            }
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RenameAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ModelContextProtocol.McpException("sheetName and targetName are required for rename action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.RenameAsync(batch, sheetName, targetName));

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
                    "Renaming multiple sheets? Keep reusing this session for best performance"
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
        string sessionId,
        string? sheetName,
        string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ModelContextProtocol.McpException("sheetName and targetName are required for copy action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.CopyAsync(batch, sheetName, targetName));

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
                    "Copying multiple sheets? Keep reusing this session for best performance"
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
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for delete action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.DeleteAsync(batch, sheetName));

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
                    "Deleting multiple sheets? Keep reusing this session for best performance"
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
        string sessionId,
        string? sheetName,
        int? red,
        int? green,
        int? blue)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for set-tab-color action");

        if (!red.HasValue)
            throw new ModelContextProtocol.McpException("red value (0-255) is required for set-tab-color action");
        if (!green.HasValue)
            throw new ModelContextProtocol.McpException("green value (0-255) is required for set-tab-color action");
        if (!blue.HasValue)
            throw new ModelContextProtocol.McpException("blue value (0-255) is required for set-tab-color action");

        // Extract values after validation (null checks above guarantee non-null)
        int redValue = red.Value;
        int greenValue = green.Value;
        int blueValue = blue.Value;

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.SetTabColorAsync(batch, sheetName, redValue, greenValue, blueValue));
        string hexColor = $"#{redValue:X2}{greenValue:X2}{blueValue:X2}";

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Tab color set to {hexColor} (RGB: {redValue}, {greenValue}, {blueValue}) for sheet '{sheetName}'."
                : $"Failed to set tab color: {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use 'get-tab-color' to verify the color was applied",
                    "Apply consistent colors to related sheets for organization",
                    "Coloring multiple sheets? Keep reusing this session for best performance"
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
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for get-tab-color action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.GetTabColorAsync(batch, sheetName));

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
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for clear-tab-color action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.ClearTabColorAsync(batch, sheetName));

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
                    "Clearing colors on multiple sheets? Keep reusing this session for best performance"
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
        string sessionId,
        string? sheetName,
        string? visibility)
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

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.SetVisibilityAsync(batch, sheetName, visibilityLevel));

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
                    "Managing visibility for multiple sheets? Keep reusing this session for best performance"
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
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for get-visibility action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.GetVisibilityAsync(batch, sheetName));

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
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for show action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.ShowAsync(batch, sheetName));

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
                    "Showing multiple sheets? Keep reusing this session for best performance"
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
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for hide action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.HideAsync(batch, sheetName));

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
                    "Hiding multiple sheets? Keep reusing this session for best performance"
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
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for very-hide action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await sheetCommands.VeryHideAsync(batch, sheetName));

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
                    "Protecting multiple sheets? Keep reusing this session for best performance"
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
