using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call
#pragma warning disable IDE0060 // batchId parameter kept for compatibility, will be removed in final cleanup phase

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
        // Use filePath-based API (ignoring batchId for now - will be removed in final cleanup)
        var result = await sheetCommands.ListAsync(excelPath);

        var count = result.Worksheets?.Count ?? 0;

        return JsonSerializer.Serialize(new
        {
            success = result.Success,
            worksheets = result.Worksheets,
            workflowHint = $"Found {count} worksheet(s). Use excel_range for data operations."
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

        // Use filePath-based API
        var result = await sheetCommands.CreateAsync(excelPath, sheetName);

        // Auto-save after create
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            workflowHint = $"Worksheet '{sheetName}' created successfully."
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

        // Use filePath-based API
        var result = await sheetCommands.RenameAsync(excelPath, sheetName, targetName);

        // Auto-save after rename
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Worksheet '{sheetName}' renamed to '{targetName}' successfully."
                : $"Failed to rename worksheet: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.CopyAsync(excelPath, sheetName, targetName);

        // Auto-save after copy
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Worksheet '{sheetName}' copied to '{targetName}' successfully."
                : $"Failed to copy worksheet: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.DeleteAsync(excelPath, sheetName);

        // Auto-save after delete
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Worksheet '{sheetName}' deleted successfully. Data is permanently removed."
                : $"Failed to delete worksheet: {result.ErrorMessage}"
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

        // Extract values after validation
        int redValue = red.Value;
        int greenValue = green.Value;
        int blueValue = blue.Value;

        // Use filePath-based API
        var result = await sheetCommands.SetTabColorAsync(excelPath, sheetName, redValue, greenValue, blueValue);

        // Auto-save after setting color
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        string hexColor = $"#{redValue:X2}{greenValue:X2}{blueValue:X2}";

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Tab color set to {hexColor} (RGB: {redValue}, {greenValue}, {blueValue}) for sheet '{sheetName}'."
                : $"Failed to set tab color: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.GetTabColorAsync(excelPath, sheetName);

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
                : $"Failed to get tab color: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.ClearTabColorAsync(excelPath, sheetName);

        // Auto-save after clearing color
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Tab color cleared for sheet '{sheetName}' (reset to default)."
                : $"Failed to clear tab color: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.SetVisibilityAsync(excelPath, sheetName, visibilityLevel);

        // Auto-save after setting visibility
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Visibility set to '{visibility}' for sheet '{sheetName}'."
                : $"Failed to set visibility: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.GetVisibilityAsync(excelPath, sheetName);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Visibility,
            result.VisibilityName,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sheet '{sheetName}' visibility is '{result.VisibilityName}'."
                : $"Failed to get visibility: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.ShowAsync(excelPath, sheetName);

        // Auto-save after showing
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sheet '{sheetName}' is now visible in the workbook."
                : $"Failed to show sheet: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.HideAsync(excelPath, sheetName);

        // Auto-save after hiding
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sheet '{sheetName}' is now hidden (users can unhide via Excel UI)."
                : $"Failed to hide sheet: {result.ErrorMessage}"
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

        // Use filePath-based API
        var result = await sheetCommands.VeryHideAsync(excelPath, sheetName);

        // Auto-save after very hiding
        if (result.Success)
        {
            await FileHandleManager.Instance.SaveAsync(excelPath);
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Sheet '{sheetName}' is now very hidden (requires code to unhide)."
                : $"Failed to very hide sheet: {result.ErrorMessage}"
        }, ExcelToolsBase.JsonOptions);
    }
}
