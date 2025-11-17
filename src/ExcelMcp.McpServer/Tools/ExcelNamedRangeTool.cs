using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel named range (parameter) operations.
/// </summary>
[McpServerToolType]
public static class ExcelNamedRangeTool
{
    // Cache JsonSerializerOptions to satisfy CA1869
    private static readonly JsonSerializerOptions s_jsonOptions = new() { PropertyNameCaseInsensitive = true };

    /// <summary>
    /// Manage Excel parameters (named ranges) - configuration values and reusable references
    /// </summary>
    [McpServerTool(Name = "excel_namedrange")]
    [Description(@"Manage Excel named ranges")]
    public static string ExcelParameter(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        NamedRangeAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

        [StringLength(255, MinimumLength = 1)]
        [Description("Named range name (for read, write, create, update, delete actions)")]
        string? namedRangeName = null,

        [Description("Named range value (for write action) or cell reference (for create/update actions, e.g., 'Sheet1!A1')")]
        string? value = null,

        [Description("JSON array of named ranges for create-bulk action: [{name: 'Name', reference: 'Sheet1!A1', value: 'text'}, ...]")]
        string? namedRangesJson = null)
    {
        try
        {
            var namedRangeCommands = new NamedRangeCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                NamedRangeAction.List => ListNamedRangesAsync(namedRangeCommands, sessionId),
                NamedRangeAction.Read => ReadNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName),
                NamedRangeAction.Write => WriteNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                NamedRangeAction.Create => CreateNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                NamedRangeAction.CreateBulk => CreateBulkNamedRangesAsync(namedRangeCommands, sessionId, namedRangesJson),
                NamedRangeAction.Update => UpdateNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName, value),
                NamedRangeAction.Delete => DeleteNamedRangeAsync(namedRangeCommands, sessionId, namedRangeName),
                _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
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

    private static string ListNamedRangesAsync(NamedRangeCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.List(batch));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        var count = result.NamedRanges?.Count ?? 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.NamedRanges
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ArgumentException("namedRangeName is required for read action", nameof(namedRangeName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Read(batch, namedRangeName));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.NamedRangeName,
            result.Value
        }, ExcelToolsBase.JsonOptions);
    }

    private static string WriteNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || value == null)
            throw new ArgumentException("namedRangeName and value are required for write action", "namedRangeName,value");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Write(batch, namedRangeName, value));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        return JsonSerializer.Serialize(new
        {
            result.Success
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ArgumentException("namedRangeName and value (cell reference) are required for update action", "namedRangeName,value");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Update(batch, namedRangeName, value));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName, string? value)
    {
        if (string.IsNullOrEmpty(namedRangeName) || string.IsNullOrEmpty(value))
            throw new ArgumentException("namedRangeName and value (cell reference) are required for create action", "namedRangeName,value");

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Create(batch, namedRangeName, value));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Add workflow hints
        return JsonSerializer.Serialize(new
        {
            result.Success
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteNamedRangeAsync(NamedRangeCommands commands, string sessionId, string? namedRangeName)
    {
        if (string.IsNullOrEmpty(namedRangeName))
            throw new ArgumentException("namedRangeName is required for delete action", nameof(namedRangeName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Delete(batch, namedRangeName));

        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success ? "Formulas referencing this named range will show #NAME? error" : null
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateBulkNamedRangesAsync(NamedRangeCommands commands, string sessionId, string? namedRangesJson)
    {
        if (string.IsNullOrWhiteSpace(namedRangesJson))
            throw new ArgumentException("namedRangesJson is required for create-bulk action", nameof(namedRangesJson));

        // Deserialize JSON array of named range definitions
        List<NamedRangeDefinition>? parameters;
        try
        {
            parameters = JsonSerializer.Deserialize<List<NamedRangeDefinition>>(
                namedRangesJson,
                s_jsonOptions);

            if (parameters == null || parameters.Count == 0)
                throw new ArgumentException("namedRangesJson must contain at least one named range definition", nameof(namedRangesJson));
        }
        catch (JsonException ex)
        {
            throw new ArgumentException($"Invalid namedRangesJson format: {ex.Message}", nameof(namedRangesJson));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CreateBulk(batch, parameters));

        // Add workflow hints (CreateBulk returns OperationResult, not specialized type)
        return JsonSerializer.Serialize(new
        {
            result.Success
        }, ExcelToolsBase.JsonOptions);
    }
}

