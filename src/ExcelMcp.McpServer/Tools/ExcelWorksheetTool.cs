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
                    : "Retry acceptable if issue is transient."
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
            worksheets = result.Worksheets
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
            result.Success
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
            result.ErrorMessage
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
            result.ErrorMessage
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
            result.ErrorMessage
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
            result.ErrorMessage
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
            result.ErrorMessage
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
            result.ErrorMessage
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
            result.ErrorMessage
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
            result.ErrorMessage
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
            result.ErrorMessage
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
            workflowHint = result.Success ? "Sheet now hidden (users can unhide via Excel UI)" : null
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
            workflowHint = result.Success ? "Sheet now very-hidden (not visible even in VBA, requires code to unhide)" : null
        }, ExcelToolsBase.JsonOptions);
    }
}
