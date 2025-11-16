using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands.QueryTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
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
    public static string ExcelQueryTable(
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
                QueryTableAction.List => ListQueryTablesAsync(queryTableCommands, sessionId),
                QueryTableAction.Read => ReadQueryTableAsync(queryTableCommands, sessionId, queryTableName),
                QueryTableAction.CreateFromConnection => CreateFromConnectionAsync(queryTableCommands, sessionId, sheetName, queryTableName, connectionName, range, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, refreshImmediately),
                QueryTableAction.CreateFromQuery => CreateFromQueryAsync(queryTableCommands, sessionId, sheetName, queryTableName, queryName, range, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, refreshImmediately),
                QueryTableAction.Refresh => RefreshQueryTableAsync(queryTableCommands, sessionId, queryTableName),
                QueryTableAction.RefreshAll => RefreshAllQueryTablesAsync(queryTableCommands, sessionId),
                QueryTableAction.UpdateProperties => UpdatePropertiesAsync(queryTableCommands, sessionId, queryTableName, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth),
                QueryTableAction.Delete => DeleteQueryTableAsync(queryTableCommands, sessionId, queryTableName),
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

    private static string ListQueryTablesAsync(QueryTableCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryTables,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadQueryTableAsync(QueryTableCommands commands, string sessionId, string? queryTableName)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ArgumentException("queryTableName is required for read action", nameof(queryTableName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Read(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryTable,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateFromConnectionAsync(
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
            throw new ArgumentException("sheetName is required for create-from-connection action", nameof(sheetName));

        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ArgumentException("queryTableName is required for create-from-connection action", nameof(queryTableName));

        if (string.IsNullOrWhiteSpace(connectionName))
            throw new ArgumentException("connectionName is required for create-from-connection action", nameof(connectionName));

        var options = new PowerQueryHelpers.QueryTableCreateOptions
        {
            Name = queryTableName,
            BackgroundQuery = backgroundQuery ?? false,
            RefreshOnFileOpen = refreshOnFileOpen ?? false,
            SavePassword = savePassword ?? false,
            PreserveColumnInfo = preserveColumnInfo ?? true,
            PreserveFormatting = preserveFormatting ?? true,
            AdjustColumnWidth = adjustColumnWidth ?? true,
            RefreshImmediately = refreshImmediately ?? true
        };

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CreateFromConnection(batch, sheetName, queryTableName, connectionName, range ?? "A1", options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateFromQueryAsync(
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
            throw new ArgumentException("sheetName is required for create-from-query action", nameof(sheetName));

        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ArgumentException("queryTableName is required for create-from-query action", nameof(queryTableName));

        if (string.IsNullOrWhiteSpace(queryName))
            throw new ArgumentException("queryName is required for create-from-query action", nameof(queryName));

        var options = new PowerQueryHelpers.QueryTableCreateOptions
        {
            Name = queryTableName,
            BackgroundQuery = backgroundQuery ?? false,
            RefreshOnFileOpen = refreshOnFileOpen ?? false,
            SavePassword = savePassword ?? false,
            PreserveColumnInfo = preserveColumnInfo ?? true,
            PreserveFormatting = preserveFormatting ?? true,
            AdjustColumnWidth = adjustColumnWidth ?? true,
            RefreshImmediately = refreshImmediately ?? true
        };

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CreateFromQuery(batch, sheetName, queryTableName, queryName, range ?? "A1", options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshQueryTableAsync(QueryTableCommands commands, string sessionId, string? queryTableName)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ArgumentException("queryTableName is required for refresh action", nameof(queryTableName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Refresh(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshAllQueryTablesAsync(QueryTableCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.RefreshAll(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UpdatePropertiesAsync(
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
            throw new ArgumentException("queryTableName is required for update-properties action", nameof(queryTableName));

        var options = new QueryTableUpdateOptions
        {
            BackgroundQuery = backgroundQuery,
            RefreshOnFileOpen = refreshOnFileOpen,
            SavePassword = savePassword,
            PreserveColumnInfo = preserveColumnInfo,
            PreserveFormatting = preserveFormatting,
            AdjustColumnWidth = adjustColumnWidth
        };

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.UpdateProperties(batch, queryTableName, options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteQueryTableAsync(QueryTableCommands commands, string sessionId, string? queryTableName)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ArgumentException("queryTableName is required for delete action", nameof(queryTableName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Delete(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}

