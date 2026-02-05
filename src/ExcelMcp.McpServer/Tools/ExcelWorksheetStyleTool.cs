using System.ComponentModel;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel worksheet styling operations (tab colors, visibility)
/// </summary>
[McpServerToolType]
public static partial class ExcelWorksheetStyleTool
{
    /// <summary>
    /// Worksheet styling operations for tab colors and visibility.
    ///
    /// TAB COLORS: Use RGB values (0-255 each) to set custom tab colors for visual organization.
    /// VISIBILITY LEVELS:
    /// - 'visible': Normal visible sheet
    /// - 'hidden': Hidden but accessible via Format > Sheet > Unhide
    /// - 'veryhidden': Only accessible via VBA (protection against casual unhiding)
    ///
    /// Related: Use excel_worksheet for sheet lifecycle (create, rename, copy, delete, move).
    /// </summary>
    /// <param name="action">The styling operation to perform</param>
    /// <param name="sessionId">Session ID from excel_file 'open'. Required for all actions.</param>
    /// <param name="sheetName">Name of the worksheet to style. Required for all actions.</param>
    /// <param name="red">Red color component (0-255). Required for: set-tab-color</param>
    /// <param name="green">Green color component (0-255). Required for: set-tab-color</param>
    /// <param name="blue">Blue color component (0-255). Required for: set-tab-color</param>
    /// <param name="visibility">Visibility level: 'visible', 'hidden', or 'veryhidden'. Required for: set-visibility</param>
    [McpServerTool(Name = "excel_worksheet_style", Title = "Excel Worksheet Style Operations", Destructive = true)]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelWorksheetStyle(
        SheetStyleAction action,
        string sessionId,
        string sheetName,
        [DefaultValue(null)] int? red,
        [DefaultValue(null)] int? green,
        [DefaultValue(null)] int? blue,
        [DefaultValue(null)] string? visibility)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_worksheet_style",
            ServiceRegistry.SheetStyle.ToActionString(action),
            () =>
            {
                return action switch
                {
                    SheetStyleAction.SetTabColor => ForwardSetTabColor(sessionId, sheetName, red, green, blue),
                    SheetStyleAction.GetTabColor => ForwardGetTabColor(sessionId, sheetName),
                    SheetStyleAction.ClearTabColor => ForwardClearTabColor(sessionId, sheetName),
                    SheetStyleAction.SetVisibility => ForwardSetVisibility(sessionId, sheetName, visibility),
                    SheetStyleAction.GetVisibility => ForwardGetVisibility(sessionId, sheetName),
                    SheetStyleAction.Show => ForwardShow(sessionId, sheetName),
                    SheetStyleAction.Hide => ForwardHide(sessionId, sheetName),
                    SheetStyleAction.VeryHide => ForwardVeryHide(sessionId, sheetName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({ServiceRegistry.SheetStyle.ToActionString(action)})", nameof(action))
                };
            });
    }

    private static string ForwardSetTabColor(string sessionId, string? sheetName, int? red, int? green, int? blue)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for set-tab-color action", nameof(sheetName));
        if (!red.HasValue)
            throw new ArgumentException("red (0-255) is required for set-tab-color action", nameof(red));
        if (!green.HasValue)
            throw new ArgumentException("green (0-255) is required for set-tab-color action", nameof(green));
        if (!blue.HasValue)
            throw new ArgumentException("blue (0-255) is required for set-tab-color action", nameof(blue));

        return ExcelToolsBase.ForwardToService("sheet.set-tab-color", sessionId, new { sheetName, red = red.Value, green = green.Value, blue = blue.Value });
    }

    private static string ForwardGetTabColor(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for get-tab-color action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.get-tab-color", sessionId, new { sheetName });
    }

    private static string ForwardClearTabColor(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for clear-tab-color action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.clear-tab-color", sessionId, new { sheetName });
    }

    private static string ForwardSetVisibility(string sessionId, string? sheetName, string? visibility)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for set-visibility action", nameof(sheetName));
        if (string.IsNullOrEmpty(visibility))
            throw new ArgumentException("visibility (visible|hidden|veryhidden) is required for set-visibility action", nameof(visibility));

        return ExcelToolsBase.ForwardToService("sheet.set-visibility", sessionId, new { sheetName, visibility });
    }

    private static string ForwardGetVisibility(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for get-visibility action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.get-visibility", sessionId, new { sheetName });
    }

    private static string ForwardShow(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for show action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.show", sessionId, new { sheetName });
    }

    private static string ForwardHide(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for hide action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.hide", sessionId, new { sheetName });
    }

    private static string ForwardVeryHide(string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for very-hide action", nameof(sheetName));

        return ExcelToolsBase.ForwardToService("sheet.very-hide", sessionId, new { sheetName });
    }
}




