using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel worksheet lifecycle and appearance (create, rename, copy, delete, tab colors, visibility).
/// </summary>
[McpServerToolType]
public static partial class ExcelWorksheetTool
{
    /// <summary>
    /// Manage Excel worksheets: lifecycle, tab colors, visibility.
    ///
    /// CROSS-FILE OPERATIONS (RECOMMENDED - no session required):
    /// - copy-to-file: Copy sheet to another file. Provide sourceFile, sheetName, targetFile.
    /// - move-to-file: Move sheet to another file. Provide sourceFile, sheetName, targetFile.
    /// These operations are ATOMIC - they handle file opening/saving/closing internally.
    ///
    /// SESSION-BASED OPERATIONS (require active session):
    /// All other operations require sessionId from excel_file 'open' action.
    ///
    /// TAB COLORS: RGB values 0-255 for red, green, blue.
    /// POSITIONING: Use beforeSheet OR afterSheet (not both). If neither, sheet goes to end.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="sessionId">Active Excel session ID (required for session-based actions, not needed for copy-to-file/move-to-file)</param>
    /// <param name="sourceFile">Source file path (required for copy-to-file and move-to-file actions)</param>
    /// <param name="sheetName">Worksheet name (required for most actions)</param>
    /// <param name="targetFile">Target file path (required for copy-to-file and move-to-file actions)</param>
    /// <param name="targetName">New sheet name (for rename, copy, or copy-to-file actions)</param>
    /// <param name="beforeSheet">Position sheet before this sheet</param>
    /// <param name="afterSheet">Position sheet after this sheet</param>
    /// <param name="red">Red component (0-255) for set-tab-color action</param>
    /// <param name="green">Green component (0-255) for set-tab-color action</param>
    /// <param name="blue">Blue component (0-255) for set-tab-color action</param>
    /// <param name="visibility">Visibility level for set-visibility action: visible (normal), hidden (user can unhide), veryhidden (requires code to unhide)</param>
    [McpServerTool(Name = "excel_worksheet", Title = "Excel Worksheet Operations")]
    [McpMeta("category", "structure")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelWorksheet(
        WorksheetAction action,
        [DefaultValue(null)] string? sessionId,
        [DefaultValue(null)] string? sourceFile,
        [DefaultValue(null)] string? sheetName,
        [DefaultValue(null)] string? targetFile,
        [DefaultValue(null)] string? targetName,
        [DefaultValue(null)] string? beforeSheet,
        [DefaultValue(null)] string? afterSheet,
        [DefaultValue(null)] int? red,
        [DefaultValue(null)] int? green,
        [DefaultValue(null)] int? blue,
        [DefaultValue(null)] string? visibility)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_worksheet",
            action.ToActionString(),
            () =>
            {
                var sheetCommands = new SheetCommands();

                // Expression switch pattern for audit compliance
                return action switch
                {
                    // Atomic cross-file operations (no session required)
                    WorksheetAction.CopyToFile => CopyToFileHandler(sheetCommands, sourceFile, sheetName, targetFile, targetName, beforeSheet, afterSheet),
                    WorksheetAction.MoveToFile => MoveToFileHandler(sheetCommands, sourceFile, sheetName, targetFile, beforeSheet, afterSheet),

                    // Session-based operations
                    WorksheetAction.List => ListAsync(sheetCommands, sessionId!),
                    WorksheetAction.Create => CreateAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.Delete => DeleteAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.Rename => RenameAsync(sheetCommands, sessionId!, sheetName, targetName),
                    WorksheetAction.Copy => CopyAsync(sheetCommands, sessionId!, sheetName, targetName),
                    WorksheetAction.Move => MoveAsync(sheetCommands, sessionId!, sheetName, beforeSheet, afterSheet),
                    WorksheetAction.SetTabColor => SetTabColorAsync(sheetCommands, sessionId!, sheetName, red, green, blue),
                    WorksheetAction.GetTabColor => GetTabColorAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.ClearTabColor => ClearTabColorAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.SetVisibility => SetVisibilityAsync(sheetCommands, sessionId!, sheetName, visibility),
                    WorksheetAction.GetVisibility => GetVisibilityAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.Show => ShowAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.Hide => HideAsync(sheetCommands, sessionId!, sheetName),
                    WorksheetAction.VeryHide => VeryHideAsync(sheetCommands, sessionId!, sheetName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Create(batch, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' created successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string RenameAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ArgumentException("sheetName and targetName are required for rename action", "sheetName,targetName");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Rename(batch, sheetName, targetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' renamed to '{targetName}' successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string CopyAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName,
        string? targetName)
    {
        if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
            throw new ArgumentException("sheetName and targetName are required for copy action", "sheetName,targetName");

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Copy(batch, sheetName, targetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' copied to '{targetName}' successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string DeleteAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for delete action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Delete(batch, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' deleted successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.SetTabColor(batch, sheetName, redValue, greenValue, blueValue);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Tab color for sheet '{sheetName}' set successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.ClearTabColor(batch, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Tab color for sheet '{sheetName}' cleared successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.SetVisibility(batch, sheetName, visibilityLevel);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' visibility set to {visibilityLevel} successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Show(batch, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' shown successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string HideAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for hide action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Hide(batch, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' hidden successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string VeryHideAsync(
        SheetCommands sheetCommands,
        string sessionId,
        string? sheetName)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for very-hide action", nameof(sheetName));

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.VeryHide(batch, sheetName);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' very-hidden successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch =>
                {
                    sheetCommands.Move(batch, sheetName, beforeSheet, afterSheet);
                    return 0;
                });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' moved successfully."
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    // === ATOMIC CROSS-FILE OPERATIONS ===

    private static string CopyToFileHandler(
        SheetCommands sheetCommands,
        string? sourceFile,
        string? sheetName,
        string? targetFile,
        string? targetName,
        string? beforeSheet,
        string? afterSheet)
    {
        if (string.IsNullOrEmpty(sourceFile))
            throw new ArgumentException("sourceFile is required for copy-to-file action", nameof(sourceFile));

        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for copy-to-file action", nameof(sheetName));

        if (string.IsNullOrEmpty(targetFile))
            throw new ArgumentException("targetFile is required for copy-to-file action", nameof(targetFile));

        try
        {
            sheetCommands.CopyToFile(sourceFile, sheetName, targetFile, targetName, beforeSheet, afterSheet);

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' copied from '{Path.GetFileName(sourceFile)}' to '{Path.GetFileName(targetFile)}' successfully.",
                sourceFile,
                targetFile,
                sheetName,
                targetSheetName = targetName ?? sheetName
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string MoveToFileHandler(
        SheetCommands sheetCommands,
        string? sourceFile,
        string? sheetName,
        string? targetFile,
        string? beforeSheet,
        string? afterSheet)
    {
        if (string.IsNullOrEmpty(sourceFile))
            throw new ArgumentException("sourceFile is required for move-to-file action", nameof(sourceFile));

        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for move-to-file action", nameof(sheetName));

        if (string.IsNullOrEmpty(targetFile))
            throw new ArgumentException("targetFile is required for move-to-file action", nameof(targetFile));

        try
        {
            sheetCommands.MoveToFile(sourceFile, sheetName, targetFile, beforeSheet, afterSheet);

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Sheet '{sheetName}' moved from '{Path.GetFileName(sourceFile)}' to '{Path.GetFileName(targetFile)}' successfully.",
                sourceFile,
                targetFile,
                sheetName
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }
}

