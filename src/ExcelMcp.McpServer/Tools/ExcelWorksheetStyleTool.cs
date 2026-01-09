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
    /// Worksheet styling - tab colors (RGB 0-255) and visibility (visible|hidden|veryhidden).
    /// Related: excel_worksheet (lifecycle)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="sid">Session ID</param>
    /// <param name="sn">Sheet name</param>
    /// <param name="r">Red 0-255</param>
    /// <param name="g">Green 0-255</param>
    /// <param name="b">Blue 0-255</param>
    /// <param name="vis">visible|hidden|veryhidden</param>
    [McpServerTool(Name = "excel_worksheet_style", Title = "Excel Worksheet Style Operations")]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelWorksheetStyle(
        WorksheetStyleAction action,
        string sid,
        string sn,
        [DefaultValue(null)] int? r,
        [DefaultValue(null)] int? g,
        [DefaultValue(null)] int? b,
        [DefaultValue(null)] string? vis)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_worksheet_style",
            action.ToActionString(),
            () =>
            {
                var sheetCommands = new SheetCommands();

                return action switch
                {
                    WorksheetStyleAction.SetTabColor => SetTabColor(sheetCommands, sid, sn, r, g, b),
                    WorksheetStyleAction.GetTabColor => GetTabColor(sheetCommands, sid, sn),
                    WorksheetStyleAction.ClearTabColor => ClearTabColor(sheetCommands, sid, sn),
                    WorksheetStyleAction.SetVisibility => SetVisibility(sheetCommands, sid, sn, vis),
                    WorksheetStyleAction.GetVisibility => GetVisibility(sheetCommands, sid, sn),
                    WorksheetStyleAction.Show => Show(sheetCommands, sid, sn),
                    WorksheetStyleAction.Hide => Hide(sheetCommands, sid, sn),
                    WorksheetStyleAction.VeryHide => VeryHide(sheetCommands, sid, sn),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string SetTabColor(SheetCommands sheetCommands, string sessionId, string? sheetName, int? red, int? green, int? blue)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sn is required for set-tab-color action", nameof(sheetName));
        if (!red.HasValue)
            throw new ArgumentException("r (0-255) is required for set-tab-color action", nameof(red));
        if (!green.HasValue)
            throw new ArgumentException("g (0-255) is required for set-tab-color action", nameof(green));
        if (!blue.HasValue)
            throw new ArgumentException("b (0-255) is required for set-tab-color action", nameof(blue));

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
            throw new ArgumentException("sn is required for get-tab-color action", nameof(sheetName));

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
            throw new ArgumentException("sn is required for clear-tab-color action", nameof(sheetName));

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
            throw new ArgumentException("sn is required for set-visibility action", nameof(sheetName));
        if (string.IsNullOrEmpty(visibility))
            throw new ArgumentException("vis (visible|hidden|veryhidden) is required for set-visibility action", nameof(visibility));

        SheetVisibility visibilityLevel = visibility.ToLowerInvariant() switch
        {
            "visible" => SheetVisibility.Visible,
            "hidden" => SheetVisibility.Hidden,
            "veryhidden" => SheetVisibility.VeryHidden,
            _ => throw new ArgumentException($"Invalid vis '{visibility}'. Use: visible, hidden, or veryhidden", nameof(visibility))
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
            throw new ArgumentException("sn is required for get-visibility action", nameof(sheetName));

        var result = ExcelToolsBase.WithSession(sessionId, batch => sheetCommands.GetVisibility(batch, sheetName));

        return JsonSerializer.Serialize(new { result.Success, result.Visibility, result.VisibilityName, result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string Show(SheetCommands sheetCommands, string sessionId, string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sn is required for show action", nameof(sheetName));

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
            throw new ArgumentException("sn is required for hide action", nameof(sheetName));

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
            throw new ArgumentException("sn is required for very-hide action", nameof(sheetName));

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
