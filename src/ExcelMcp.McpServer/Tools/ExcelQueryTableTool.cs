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
        bool? refreshImmediately = null,

        [Description("Optional batch ID for grouping operations")]
        string? batchId = null)
    {
        try
        {
            var queryTableCommands = new QueryTableCommands();

            return action switch
            {
                QueryTableAction.List => await ListQueryTablesAsync(queryTableCommands, excelPath, batchId),
                QueryTableAction.Get => await GetQueryTableAsync(queryTableCommands, excelPath, queryTableName, batchId),
                QueryTableAction.CreateFromConnection => await CreateFromConnectionAsync(queryTableCommands, excelPath, sheetName, queryTableName, connectionName, range, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, refreshImmediately, batchId),
                QueryTableAction.CreateFromQuery => await CreateFromQueryAsync(queryTableCommands, excelPath, sheetName, queryTableName, queryName, range, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, refreshImmediately, batchId),
                QueryTableAction.Refresh => await RefreshQueryTableAsync(queryTableCommands, excelPath, queryTableName, batchId),
                QueryTableAction.RefreshAll => await RefreshAllQueryTablesAsync(queryTableCommands, excelPath, batchId),
                QueryTableAction.UpdateProperties => await UpdatePropertiesAsync(queryTableCommands, excelPath, queryTableName, backgroundQuery, refreshOnFileOpen, savePassword, preserveColumnInfo, preserveFormatting, adjustColumnWidth, batchId),
                QueryTableAction.Delete => await DeleteQueryTableAsync(queryTableCommands, excelPath, queryTableName, batchId),
                _ => throw new ModelContextProtocol.McpException($"Unknown action: {action}")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw;
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw;
        }
    }

    private static async Task<string> ListQueryTablesAsync(QueryTableCommands commands, string excelPath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.ListAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryTables,
            result.ErrorMessage,
            workflowHint = result.Success
                ? result.QueryTables?.Count > 0
                    ? "QueryTables listed successfully. Use 'get' to inspect details, 'refresh' to reload data, or 'update-properties' to modify settings."
                    : "No QueryTables found. Use 'create-from-connection' or 'create-from-query' to import data."
                : "Failed to list QueryTables. Check file path and ensure workbook contains valid data connections.",
            suggestedNextActions = result.Success
                ? result.QueryTables?.Count > 0
                    ? new[] { "Use get action to view QueryTable details", "Use refresh action to reload data", "Use update-properties to modify refresh settings" }
                    : ["Use create-from-connection to import from data connection", "Use create-from-query to import from Power Query"]
                : ["Verify file path is correct", "Check if workbook has data connections", "Review error message for details"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetQueryTableAsync(QueryTableCommands commands, string excelPath, string? queryTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for get action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: false,
            async (batch) => await commands.GetAsync(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.QueryTable,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"QueryTable '{queryTableName}' details retrieved. Review properties to understand data source and refresh behavior."
                : $"Failed to get QueryTable '{queryTableName}'. Verify name is correct and QueryTable exists in workbook.",
            suggestedNextActions = result.Success
                ? new[] { "Use refresh action to reload data", "Use update-properties to modify settings", "Use delete to remove QueryTable" }
                : ["Use list action to see all available QueryTables", "Check QueryTable name spelling", "Review error message for details"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateFromConnectionAsync(
        QueryTableCommands commands,
        string excelPath,
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
        bool? refreshImmediately,
        string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.CreateFromConnectionAsync(batch, sheetName, queryTableName, connectionName, range ?? "A1", options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"QueryTable '{queryTableName}' created successfully from connection '{connectionName}'. Data loaded to '{sheetName}!{range ?? "A1"}'."
                : $"Failed to create QueryTable from connection. Check connection name '{connectionName}' exists and destination sheet '{sheetName}' is valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use get action to view QueryTable properties", "Use refresh to reload data", "Use update-properties to modify settings" }
                : ["Use excel_connection list to verify connection exists", "Check sheet name is valid", "Review error message for details"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateFromQueryAsync(
        QueryTableCommands commands,
        string excelPath,
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
        bool? refreshImmediately,
        string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.CreateFromQueryAsync(batch, sheetName, queryTableName, queryName, range ?? "A1", options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"QueryTable '{queryTableName}' created successfully from Power Query '{queryName}'. Data loaded to '{sheetName}!{range ?? "A1"}'."
                : $"Failed to create QueryTable from Power Query. Check query name '{queryName}' exists and destination sheet '{sheetName}' is valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use get action to view QueryTable properties", "Use refresh to reload data", "Use update-properties to modify settings" }
                : ["Use excel_powerquery list to verify query exists", "Check sheet name is valid", "Review error message for details"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshQueryTableAsync(QueryTableCommands commands, string excelPath, string? queryTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for refresh action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.RefreshAsync(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"QueryTable '{queryTableName}' refreshed successfully. Data reloaded from source."
                : $"Failed to refresh QueryTable '{queryTableName}'. Check data source connectivity and query validity.",
            suggestedNextActions = result.Success
                ? new[] { "Use get action to verify updated data", "Use update-properties to modify refresh settings", "Review refreshed worksheet" }
                : ["Verify data source is accessible", "Check connection credentials", "Review error message for connectivity issues"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshAllQueryTablesAsync(QueryTableCommands commands, string excelPath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.RefreshAllAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? "All QueryTables refreshed successfully. All data reloaded from sources."
                : "Failed to refresh all QueryTables. One or more data sources may be inaccessible.",
            suggestedNextActions = result.Success
                ? new[] { "Use list to see all QueryTables", "Use get to inspect individual QueryTables", "Review workbook for updated data" }
                : ["Use refresh action to refresh individual QueryTables", "Check data source connectivity", "Review error message for specific failures"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdatePropertiesAsync(
        QueryTableCommands commands,
        string excelPath,
        string? queryTableName,
        bool? backgroundQuery,
        bool? refreshOnFileOpen,
        bool? savePassword,
        bool? preserveColumnInfo,
        bool? preserveFormatting,
        bool? adjustColumnWidth,
        string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.UpdatePropertiesAsync(batch, queryTableName, options));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"QueryTable '{queryTableName}' properties updated successfully. New settings will apply on next refresh."
                : $"Failed to update QueryTable '{queryTableName}' properties. Verify QueryTable exists and property values are valid.",
            suggestedNextActions = result.Success
                ? new[] { "Use get action to verify updated properties", "Use refresh to test new settings", "Review QueryTable behavior" }
                : ["Use list to verify QueryTable exists", "Check property values are valid", "Review error message for details"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteQueryTableAsync(QueryTableCommands commands, string excelPath, string? queryTableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(queryTableName))
            throw new ModelContextProtocol.McpException("queryTableName is required for delete action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            excelPath,
            save: true,
            async (batch) => await commands.DeleteAsync(batch, queryTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"QueryTable '{queryTableName}' deleted successfully. QueryTable removed from workbook."
                : $"Failed to delete QueryTable '{queryTableName}'. Verify QueryTable exists and is not protected.",
            suggestedNextActions = result.Success
                ? new[] { "Use list to verify deletion", "Review workbook structure", "Clean up unused connections if needed" }
                : ["Use list to verify QueryTable exists", "Check if worksheet is protected", "Review error message for details"]
        }, ExcelToolsBase.JsonOptions);
    }
}
