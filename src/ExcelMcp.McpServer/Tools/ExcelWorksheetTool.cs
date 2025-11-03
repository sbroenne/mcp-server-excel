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

            // Switch directly on enum with inline logic
            switch (action)
            {
                case WorksheetAction.List:
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
                            : new[]
                            {
                                "Use excel_range for data operations (get-values, set-values, clear-*)",
                                "Use 'create' to add new worksheets",
                                "Use 'set-tab-color' to organize sheets visually",
                                inBatch ? "Continue batch operations" : count > 3 ? "Use excel_batch for multiple sheet operations (faster)" : "Use 'rename' or 'copy' to manage sheets"
                            }
                    }, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.Create:
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

                case WorksheetAction.Rename:
                {
                    if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
                        throw new ModelContextProtocol.McpException("sheetName and targetName are required for rename action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await sheetCommands.RenameAsync(batch, sheetName, targetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.Copy:
                {
                    if (string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(targetName))
                        throw new ModelContextProtocol.McpException("sheetName and targetName are required for copy action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await sheetCommands.CopyAsync(batch, sheetName, targetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.Delete:
                {
                    if (string.IsNullOrEmpty(sheetName))
                        throw new ModelContextProtocol.McpException("sheetName is required for delete action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await sheetCommands.DeleteAsync(batch, sheetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.SetTabColor:
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

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.GetTabColor:
                {
                    if (string.IsNullOrEmpty(sheetName))
                        throw new ModelContextProtocol.McpException("sheetName is required for get-tab-color action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: false,
                        async (batch) => await sheetCommands.GetTabColorAsync(batch, sheetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.ClearTabColor:
                {
                    if (string.IsNullOrEmpty(sheetName))
                        throw new ModelContextProtocol.McpException("sheetName is required for clear-tab-color action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await sheetCommands.ClearTabColorAsync(batch, sheetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.SetVisibility:
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

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.GetVisibility:
                {
                    if (string.IsNullOrEmpty(sheetName))
                        throw new ModelContextProtocol.McpException("sheetName is required for get-visibility action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: false,
                        async (batch) => await sheetCommands.GetVisibilityAsync(batch, sheetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.Show:
                {
                    if (string.IsNullOrEmpty(sheetName))
                        throw new ModelContextProtocol.McpException("sheetName is required for show action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await sheetCommands.ShowAsync(batch, sheetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.Hide:
                {
                    if (string.IsNullOrEmpty(sheetName))
                        throw new ModelContextProtocol.McpException("sheetName is required for hide action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await sheetCommands.HideAsync(batch, sheetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                case WorksheetAction.VeryHide:
                {
                    if (string.IsNullOrEmpty(sheetName))
                        throw new ModelContextProtocol.McpException("sheetName is required for very-hide action");

                    var result = await ExcelToolsBase.WithBatchAsync(
                        batchId,
                        excelPath,
                        save: true,
                        async (batch) => await sheetCommands.VeryHideAsync(batch, sheetName));

                    return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
                }

                default:
                    throw new ModelContextProtocol.McpException(
                        $"Unknown action: {action} ({action.ToActionString()})");
            }
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
}
