using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel hyperlink manipulation tool for MCP server.
/// Handles hyperlink operations for linking cells to URLs and files.
///
/// LLM Usage Patterns:
/// - Use "add" to create hyperlinks in cells
/// - Use "remove" to delete hyperlinks from cells
/// - Use "list" to view all hyperlinks in a worksheet
/// - Use "get" to inspect a specific hyperlink's properties
///
/// Note: Hyperlinks can point to URLs (http://, https://, mailto:) or file paths.
/// For web URLs, use full URL format. For files, use absolute paths.
/// </summary>
[McpServerToolType]
public static class HyperlinkTool
{
    /// <summary>
    /// Manage Excel hyperlinks - add, remove, list, and inspect hyperlinks in worksheets
    /// </summary>
    [McpServerTool(Name = "excel_hyperlink")]
    [Description("Manage Excel hyperlinks. Supports: add, remove, list, get. Optional batchId for batch sessions.")]
    public static async Task<string> Hyperlink(
        [Required]
        [RegularExpression("^(add|remove|list|get)$")]
        [Description("Action: add, remove, list, get")]
        string action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Worksheet name")]
        string sheetName,

        [RegularExpression(@"^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$")]
        [Description("Cell address (e.g., 'A1') or range (e.g., 'A1:B2') - required for add/remove/get")]
        string? cellAddress = null,

        [Description("URL or file path to link to - required for add action")]
        string? url = null,

        [StringLength(255)]
        [Description("Display text for hyperlink - optional for add action")]
        string? displayText = null,

        [StringLength(255)]
        [Description("Tooltip text shown on hover - optional for add action")]
        string? tooltip = null,

        [Description("Optional batch session ID from begin_excel_batch (for multi-operation workflows)")]
        string? batchId = null)
    {
        try
        {
            var hyperlinkCommands = new HyperlinkCommands();

            switch (action.ToLowerInvariant())
            {
                case "add":
                    return await AddHyperlink(hyperlinkCommands, excelPath, sheetName, cellAddress, url, displayText, tooltip, batchId);
                case "remove":
                    return await RemoveHyperlink(hyperlinkCommands, excelPath, sheetName, cellAddress, batchId);
                case "list":
                    return await ListHyperlinks(hyperlinkCommands, excelPath, sheetName, batchId);
                case "get":
                    return await GetHyperlink(hyperlinkCommands, excelPath, sheetName, cellAddress, batchId);
                default:
                    ExcelToolsBase.ThrowUnknownAction(action, "add", "remove", "list", "get");
                    throw new InvalidOperationException(); // Never reached
            }
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action, excelPath);
            throw;
        }
    }

    private static async Task<string> AddHyperlink(HyperlinkCommands commands, string excelPath, string sheetName, string? cellAddress, string? url, string? displayText, string? tooltip, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(cellAddress)) ExcelToolsBase.ThrowMissingParameter(nameof(cellAddress), "add");
        if (string.IsNullOrWhiteSpace(url)) ExcelToolsBase.ThrowMissingParameter(nameof(url), "add");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.AddHyperlinkAsync(batch, sheetName, cellAddress!, url!, displayText, tooltip));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"add failed for cell '{cellAddress}' in sheet '{sheetName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RemoveHyperlink(HyperlinkCommands commands, string excelPath, string sheetName, string? cellAddress, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(cellAddress)) ExcelToolsBase.ThrowMissingParameter(nameof(cellAddress), "remove");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.RemoveHyperlinkAsync(batch, sheetName, cellAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"remove failed for cell '{cellAddress}' in sheet '{sheetName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListHyperlinks(HyperlinkCommands commands, string excelPath, string sheetName, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ListHyperlinksAsync(batch, sheetName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"list failed for sheet '{sheetName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetHyperlink(HyperlinkCommands commands, string excelPath, string sheetName, string? cellAddress, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(cellAddress)) ExcelToolsBase.ThrowMissingParameter(nameof(cellAddress), "get");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.GetHyperlinkAsync(batch, sheetName, cellAddress!));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"get failed for cell '{cellAddress}' in sheet '{sheetName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
