using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.QueryTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - Workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel QueryTable operations - simple data imports with reliable persistence.
/// </summary>
[McpServerToolType]
public static class ExcelQueryTableTool
{
    /// <summary>
    /// Manage Excel QueryTables - simple data imports with reliable refresh patterns
    /// </summary>
    [McpServerTool(Name = "excel_querytable")]
    [Description(@"Manage Excel QueryTables")]
    public static async Task<string> ExcelQueryTable(
        [Required]
        [Description("Action to perform")]
        QueryTableAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

        [StringLength(255, MinimumLength = 1)]
        [Description("QueryTable name")]
        string? queryTableName = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Worksheet name")]
        string? sheetName = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Connection name (for create-from-connection)")]
        string? connectionName = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Power Query name (for create-from-query)")]
        string? queryName = null,

        [Description("Range address (default: A1)")]
        string? range = null,

        [Description("Background query setting")]
        bool? backgroundQuery = null,

        [Description("Refresh on file open setting")]
        bool? refreshOnFileOpen = null,

        [Description("Save password setting")]
        bool? savePassword = null,

        [Description("Preserve column info setting")]
        bool? preserveColumnInfo = null,

        [Description("Preserve formatting setting")]
        bool? preserveFormatting = null,

        [Description("Adjust column width setting")]
        bool? adjustColumnWidth = null,

        [Description("Refresh immediately after creation")]
        bool? refreshImmediately = null)
    {
        try
        {
            var queryTableCommands = new QueryTableCommands();

            return action switch
            {
                QueryTableAction.List => await ListQueryTablesAsync(queryTableCommands, sessionId),
                QueryTableAction.Get => await GetQueryTableAsync(queryTableCommands, sessionId, queryTableName),
                QueryTableAction.CreateFromConnection => await CreateFromConnectionAsync(queryTableCommands, sessionId, sheetName, queryTableName, connectionName, range, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, refreshImmediately),
                QueryTableAction.CreateFromQuery => await CreateFromQueryAsync(queryTableCommands, sessionId, sheetName, queryTableName, queryName, range, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, refreshImmediately),
                QueryTableAction.Refresh => await RefreshQueryTableAsync(queryTableCommands, sessionId, queryTableName),
                QueryTableAction.RefreshAll => await RefreshAllQueryTablesAsync(queryTableCommands, sessionId),
                QueryTableAction.UpdateProperties => await UpdatePropertiesAsync(queryTableCommands, sessionId, queryTableName, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth),
                QueryTableAction.Delete => await DeleteQueryTableAsync(queryTableCommands, sessionId, queryTableName),
                _ => throw new ArgumentException($"Unknown action: {action}", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed for '{excelPath}': {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static async Task<string> ListQueryTablesAsync(QueryTableCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryTables,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetQueryTableAsync(QueryTableCommands commands, string sessionId, string? queryTableName)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for get action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetAsync(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryTable,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateFromConnectionAsync(
        QueryTableCommands commands,
        string sessionId,
        string? sheetName,
        string? queryTableName,
        string? connectionName,
        string? range,
        bool? backgroundQuery,
        bool? refreshOnFileOpen,
        bool? savePassword,
        bool? preserveColumnInfo,
        bool? preserveFormatting,
        bool? adjustColumnWidth,
        bool? refreshImmediately)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for create-from-connection action");

        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for create-from-connection action");

        if (string.IsNullOrWhiteSpace(connectionName))
            throw new ModelContextProtocol.McpException("connectionName is required for create-from-connection action");

        var options = new QueryTableCreateOptions
        {
            BackgroundQuery = backgroundQuery ?? false,
            RefreshOnFileOpen = refreshOnFileOpen ?? false,
            SavePassword = savePassword ?? false,
            PreserveColumnInfo = preserveColumnInfo ?? true,
            PreserveFormatting = preserveFormatting ?? true,
            AdjustColumnWidth = adjustColumnWidth ?? true,
            RefreshImmediately = refreshImmediately ?? true
        };

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CreateFromConnectionAsync(batch, sheetName, queryTableName, connectionName, range ?? "A1", options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateFromQueryAsync(
        QueryTableCommands commands,
        string sessionId,
        string? sheetName,
        string? queryTableName,
        string? queryName,
        string? range,
        bool? backgroundQuery,
        bool? refreshOnFileOpen,
        bool? savePassword,
        bool? preserveColumnInfo,
        bool? preserveFormatting,
        bool? adjustColumnWidth,
        bool? refreshImmediately)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ModelContextProtocol.McpException("sheetName is required for create-from-query action");

        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for create-from-query action");

        if (string.IsNullOrWhiteSpace(queryName))
            throw new ModelContextProtocol.McpException("queryName is required for create-from-query action");

        var options = new QueryTableCreateOptions
        {
            BackgroundQuery = backgroundQuery ?? false,
            RefreshOnFileOpen = refreshOnFileOpen ?? false,
            SavePassword = savePassword ?? false,
            PreserveColumnInfo = preserveColumnInfo ?? true,
            PreserveFormatting = preserveFormatting ?? true,
            AdjustColumnWidth = adjustColumnWidth ?? true,
            RefreshImmediately = refreshImmediately ?? true
        };

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CreateFromQueryAsync(batch, sheetName, queryTableName, queryName, range ?? "A1", options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshQueryTableAsync(QueryTableCommands commands, string sessionId, string? queryTableName)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for refresh action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.RefreshAsync(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshAllQueryTablesAsync(QueryTableCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.RefreshAllAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdatePropertiesAsync(
        QueryTableCommands commands,
        string sessionId,
        string? queryTableName,
        bool? backgroundQuery,
        bool? refreshOnFileOpen,
        bool? savePassword,
        bool? preserveColumnInfo,
        bool? preserveFormatting,
        bool? adjustColumnWidth)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for update-properties action");

        var options = new QueryTableUpdateOptions
        {
            BackgroundQuery = backgroundQuery,
            RefreshOnFileOpen = refreshOnFileOpen,
            SavePassword = savePassword,
            PreserveColumnInfo = preserveColumnInfo,
            PreserveFormatting = preserveFormatting,
            AdjustColumnWidth = adjustColumnWidth
        };

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.UpdatePropertiesAsync(batch, queryTableName, options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteQueryTableAsync(QueryTableCommands commands, string sessionId, string? queryTableName)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for delete action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteAsync(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}
