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

POSITIONING (move, copy-to-workbook, move-to-workbook):
- Use beforeSheet OR afterSheet (not both) to specify relative position
- If neither specified, sheet is positioned at the end
")]
    public static string ExcelWorksheet(
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
        [Description("New sheet name (for rename) or target sheet name (for copy/copy-to-workbook)")]
        string? targetName = null,

        [Description("Target workbook session ID (for copy-to-workbook and move-to-workbook actions)")]
        string? targetSessionId = null,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Position sheet before this sheet (for move, copy-to-workbook, move-to-workbook)")]
        string? beforeSheet = null,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Position sheet after this sheet (for move, copy-to-workbook, move-to-workbook)")]
        string? afterSheet = null,

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
                WorksheetAction.List => ListAsync(sheetCommands, sessionId),
                WorksheetAction.Create => CreateAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.Delete => DeleteAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.Rename => RenameAsync(sheetCommands, sessionId, sheetName, targetName),
                WorksheetAction.Copy => CopyAsync(sheetCommands, sessionId, sheetName, targetName),
                WorksheetAction.Move => MoveAsync(sheetCommands, sessionId, sheetName, beforeSheet, afterSheet),
                WorksheetAction.CopyToWorkbook => CopyToWorkbookAsync(sheetCommands, sessionId, sheetName, targetSessionId, targetName, beforeSheet, afterSheet),
                WorksheetAction.MoveToWorkbook => MoveToWorkbookAsync(sheetCommands, sessionId, sheetName, targetSessionId, beforeSheet, afterSheet),
                WorksheetAction.SetTabColor => SetTabColorAsync(sheetCommands, sessionId, sheetName, red, green, blue),
                WorksheetAction.GetTabColor => GetTabColorAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.ClearTabColor => ClearTabColorAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.SetVisibility => SetVisibilityAsync(sheetCommands, sessionId, sheetName, visibility),
                WorksheetAction.GetVisibility => GetVisibilityAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.Show => ShowAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.Hide => HideAsync(sheetCommands, sessionId, sheetName),
                WorksheetAction.VeryHide => VeryHideAsync(sheetCommands, sessionId, sheetName),
                _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed: {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    // === PRIVATE HELPER METHODS ===

    private static string ListAsync(
        SheetCommands sheetCommands,
        string sessionId)
    {
        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.List(batch));

        return JsonSerializer.Serialize(new
        {
            success = result.Success,
            worksheets = result.Worksheets,
            errorMessage = result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for create action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.Create(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RenameAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ArgumentException("sheetName and targetName are required for rename action", "sheetName,targetName");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.Rename(batch, sheetName, targetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CopyAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ArgumentException("sheetName and targetName are required for copy action", "sheetName,targetName");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.Copy(batch, sheetName, targetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for delete action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.Delete(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetTabColorAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        int? red,
        int? green,
        int? blue)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for set-tab-color action", nameof(sheetName));

        if (!red.HasValue)
            throw new ArgumentException("red value (0-255) is required for set-tab-color action", nameof(red));
        if (!green.HasValue)
            throw new ArgumentException("green value (0-255) is required for set-tab-color action", nameof(green));
        if (!blue.HasValue)
            throw new ArgumentException("blue value (0-255) is required for set-tab-color action", nameof(blue));

        // Extract values after validation (null checks above guarantee non-null)
        int redValue = red.Value;
        int greenValue = green.Value;
        int blueValue = blue.Value;

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.SetTabColor(batch, sheetName, redValue, greenValue, blueValue));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetTabColorAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for get-tab-color action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.GetTabColor(batch, sheetName));

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

    private static string ClearTabColorAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for clear-tab-color action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.ClearTabColor(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string SetVisibilityAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? visibility)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for set-visibility action", nameof(sheetName));

        if (string.IsNullOrEmpty(visibility))
            throw new ArgumentException("visibility (visible|hidden|veryhidden) is required for set-visibility action", nameof(visibility));

        SheetVisibility visibilityLevel = visibility.ToLowerInvariant() switch
        {
            "visible" => SheetVisibility.Visible,
            "hidden" => SheetVisibility.Hidden,
            "veryhidden" => SheetVisibility.VeryHidden,
            _ => throw new ArgumentException($"Invalid visibility value '{visibility}'. Use: visible, hidden, or veryhidden", nameof(visibility))
        };

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.SetVisibility(batch, sheetName, visibilityLevel));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string GetVisibilityAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for get-visibility action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.GetVisibility(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Visibility,
            result.VisibilityName,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ShowAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for show action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.Show(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string HideAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for hide action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.Hide(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success ? "Sheet now hidden (users can unhide via Excel UI)" : null
        }, ExcelToolsBase.JsonOptions);
    }

    private static string VeryHideAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for very-hide action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.VeryHide(batch, sheetName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success ? "Sheet now very-hidden (not visible even in VBA, requires code to unhide)" : null
        }, ExcelToolsBase.JsonOptions);
    }

    private static string MoveAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? beforeSheet,
        string? afterSheet)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for move action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => sheetCommands.Move(batch, sheetName, beforeSheet, afterSheet));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CopyToWorkbookAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetSessionId,
        string? targetName,
        string? beforeSheet,
        string? afterSheet)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for copy-to-workbook action", nameof(sheetName));

        if (string.IsNullOrEmpty(targetSessionId))
            throw new ArgumentException("targetSessionId is required for copy-to-workbook action", nameof(targetSessionId));

        // Resolve both sessions
        var sessionManager = ExcelToolsBase.GetSessionManager();
        var sourceBatch = sessionManager.GetSession(sessionId);
        var targetBatch = sessionManager.GetSession(targetSessionId);

        if (sourceBatch == null)
            throw new InvalidOperationException($"Source session '{sessionId}' not found");

        if (targetBatch == null)
            throw new InvalidOperationException($"Target session '{targetSessionId}' not found");

        var result = sheetCommands.CopyToWorkbook(sourceBatch, sheetName, targetBatch, targetName, beforeSheet, afterSheet);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string MoveToWorkbookAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetSessionId,
        string? beforeSheet,
        string? afterSheet)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for move-to-workbook action", nameof(sheetName));

        if (string.IsNullOrEmpty(targetSessionId))
            throw new ArgumentException("targetSessionId is required for move-to-workbook action", nameof(targetSessionId));

        // Resolve both sessions
        var sessionManager = ExcelToolsBase.GetSessionManager();
        var sourceBatch = sessionManager.GetSession(sessionId);
        var targetBatch = sessionManager.GetSession(targetSessionId);

        if (sourceBatch == null)
            throw new InvalidOperationException($"Source session '{sessionId}' not found");

        if (targetBatch == null)
            throw new InvalidOperationException($"Target session '{targetSessionId}' not found");

        var result = sheetCommands.MoveToWorkbook(sourceBatch, sheetName, targetBatch, beforeSheet, afterSheet);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}

