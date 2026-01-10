using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

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
    [McpServerTool(Name = "excel_worksheet_style", Title = "Excel Worksheet Style Operations")]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelWorksheetStyle(
        WorksheetStyleAction action,
        string sessionId,
        string sheetName,
        [DefaultValue(null)] int? red,
        [DefaultValue(null)] int? green,
        [DefaultValue(null)] int? blue,
        [DefaultValue(null)] string? visibility)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_worksheet_style",
            action.ToActionString(),
            () =>
            {
                var sheetCommands = new SheetCommands();

                return action switch
                {
                    WorksheetStyleAction.SetTabColor => SetTabColor(sheetCommands, sessionId, sheetName, red, green, blue),
                    WorksheetStyleAction.GetTabColor => GetTabColor(sheetCommands, sessionId, sheetName),
                    WorksheetStyleAction.ClearTabColor => ClearTabColor(sheetCommands, sessionId, sheetName),
                    WorksheetStyleAction.SetVisibility => SetVisibility(sheetCommands, sessionId, sheetName, visibility),
                    WorksheetStyleAction.GetVisibility => GetVisibility(sheetCommands, sessionId, sheetName),
                    WorksheetStyleAction.Show => Show(sheetCommands, sessionId, sheetName),
                    WorksheetStyleAction.Hide => Hide(sheetCommands, sessionId, sheetName),
                    WorksheetStyleAction.VeryHide => VeryHide(sheetCommands, sessionId, sheetName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string SetTabColor(SheetCommands sheetCommands, string sessionId, string? sheetName, int? red, int? green, int? blue)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for set-tab-color action", nameof(sheetName));
        if (!red.HasValue)
            throw new ArgumentException("red (0-255) is required for set-tab-color action", nameof(red));
        if (!green.HasValue)
            throw new ArgumentException("green (0-255) is required for set-tab-color action", nameof(green));
        if (!blue.HasValue)
            throw new ArgumentException("blue (0-255) is required for set-tab-color action", nameof(blue));

        try
        {
            ExcelToolsBase.WithSession(sessionId, batch =>
            {
                sheetCommands.SetTabColor(batch, sheetName, red.Value, green.Value, blue.Value);
                return 0;
            });

            return JsonSerializer.Serialize(new { success = true, message = $"Tab color for sheet '{sheetName}' set successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetTabColor(SheetCommands sheetCommands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for get-tab-color action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(sessionId, batch => sheetCommands.GetTabColor(batch, sheetName));

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

    private static string ClearTabColor(SheetCommands sheetCommands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for clear-tab-color action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(sessionId, batch =>
            {
                sheetCommands.ClearTabColor(batch, sheetName);
                return 0;
            });

            return JsonSerializer.Serialize(new { success = true, message = $"Tab color for sheet '{sheetName}' cleared successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string SetVisibility(SheetCommands sheetCommands, string sessionId, string? sheetName, string? visibility)
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
            _ => throw new ArgumentException($"Invalid visibility '{visibility}'. Use: visible, hidden, or veryhidden", nameof(visibility))
        };

        try
        {
            ExcelToolsBase.WithSession(sessionId, batch =>
            {
                sheetCommands.SetVisibility(batch, sheetName, visibilityLevel);
                return 0;
            });

            return JsonSerializer.Serialize(new { success = true, message = $"Sheet '{sheetName}' visibility set to {visibilityLevel} successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string GetVisibility(SheetCommands sheetCommands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for get-visibility action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(sessionId, batch => sheetCommands.GetVisibility(batch, sheetName));

        return JsonSerializer.Serialize(new { result.Success, result.Visibility, result.VisibilityName, result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string Show(SheetCommands sheetCommands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for show action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(sessionId, batch =>
            {
                sheetCommands.Show(batch, sheetName);
                return 0;
            });

            return JsonSerializer.Serialize(new { success = true, message = $"Sheet '{sheetName}' shown successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string Hide(SheetCommands sheetCommands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for hide action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(sessionId, batch =>
            {
                sheetCommands.Hide(batch, sheetName);
                return 0;
            });

            return JsonSerializer.Serialize(new { success = true, message = $"Sheet '{sheetName}' hidden successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string VeryHide(SheetCommands sheetCommands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for very-hide action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(sessionId, batch =>
            {
                sheetCommands.VeryHide(batch, sheetName);
                return 0;
            });

            return JsonSerializer.Serialize(new { success = true, message = $"Sheet '{sheetName}' very-hidden successfully." }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }
}
