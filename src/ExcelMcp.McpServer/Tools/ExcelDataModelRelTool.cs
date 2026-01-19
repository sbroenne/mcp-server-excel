using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Data Model relationship operations.
/// Use excel_datamodel for tables and DAX measures.
/// </summary>
[McpServerToolType]
public static partial class ExcelDataModelRelTool
{
    /// <summary>
    /// Data Model relationships - link tables for cross-table DAX calculations.
    ///
    /// RELATIONSHIP REQUIREMENTS:
    /// - Both tables must exist in the Data Model first
    /// - Columns must have compatible data types
    /// - From column is typically the many-side (detail table)
    /// - To column is typically the one-side (lookup table)
    ///
    /// ACTIVE VS INACTIVE:
    /// - Only ONE active relationship can exist between two tables
    /// - Use isActive=false when creating alternative paths
    /// - DAX USERELATIONSHIP() activates inactive relationships
    ///
    /// TIMEOUT: Operations auto-timeout after 2 minutes for large Data Models.
    ///
    /// RELATED TOOLS:
    /// - excel_datamodel: Tables, measures, Data Model info
    /// - excel_table: Add tables to Data Model (add-to-datamodel action)
    /// </summary>
    /// <param name="action">The relationship operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="fromTableName">Source (many-side) table name containing the foreign key</param>
    /// <param name="fromColumnName">Column in fromTable that links to toTable</param>
    /// <param name="toTableName">Target (one-side/lookup) table name containing the primary key</param>
    /// <param name="toColumnName">Column in toTable that fromColumn links to (usually primary key)</param>
    /// <param name="isActive">True for active relationship (default), false for inactive (use with DAX USERELATIONSHIP)</param>
    [McpServerTool(Name = "excel_datamodel_rel", Title = "Excel Data Model Relationship Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelDataModelRel(
        DataModelRelAction action,
        string sessionId,
        [DefaultValue(null)] string? fromTableName,
        [DefaultValue(null)] string? fromColumnName,
        [DefaultValue(null)] string? toTableName,
        [DefaultValue(null)] string? toColumnName,
        [DefaultValue(null)] bool? isActive)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_datamodel_rel",
            action.ToActionString(),
            () =>
            {
                var commands = new DataModelCommands();

                return action switch
                {
                    DataModelRelAction.ListRelationships => ListRelationshipsAction(commands, sessionId),
                    DataModelRelAction.ReadRelationship => ReadRelationshipAction(commands, sessionId, fromTableName, fromColumnName, toTableName, toColumnName),
                    DataModelRelAction.CreateRelationship => CreateRelationshipAction(commands, sessionId, fromTableName, fromColumnName, toTableName, toColumnName, isActive),
                    DataModelRelAction.UpdateRelationship => UpdateRelationshipAction(commands, sessionId, fromTableName, fromColumnName, toTableName, toColumnName, isActive),
                    DataModelRelAction.DeleteRelationship => DeleteRelationshipAction(commands, sessionId, fromTableName, fromColumnName, toTableName, toColumnName),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListRelationshipsAction(DataModelCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.ListRelationships(batch));

        return JsonSerializer.Serialize(new { result.Success, result.Relationships, result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadRelationshipAction(DataModelCommands commands, string sessionId, string? fromTableName, string? fromColumnName, string? toTableName, string? toColumnName)
    {
        if (string.IsNullOrWhiteSpace(fromTableName))
            throw new ArgumentException("fromTableName is required for read-relationship action", nameof(fromTableName));
        if (string.IsNullOrWhiteSpace(fromColumnName))
            throw new ArgumentException("fromColumnName is required for read-relationship action", nameof(fromColumnName));
        if (string.IsNullOrWhiteSpace(toTableName))
            throw new ArgumentException("toTableName is required for read-relationship action", nameof(toTableName));
        if (string.IsNullOrWhiteSpace(toColumnName))
            throw new ArgumentException("toColumnName is required for read-relationship action", nameof(toColumnName));

        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.ReadRelationship(batch, fromTableName, fromColumnName, toTableName, toColumnName));

        return JsonSerializer.Serialize(new { result.Success, result.Relationship, result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateRelationshipAction(DataModelCommands commands, string sessionId, string? fromTableName, string? fromColumnName, string? toTableName, string? toColumnName, bool? isActive)
    {
        if (string.IsNullOrWhiteSpace(fromTableName))
            throw new ArgumentException("fromTableName is required for create-relationship action", nameof(fromTableName));
        if (string.IsNullOrWhiteSpace(fromColumnName))
            throw new ArgumentException("fromColumnName is required for create-relationship action", nameof(fromColumnName));
        if (string.IsNullOrWhiteSpace(toTableName))
            throw new ArgumentException("toTableName is required for create-relationship action", nameof(toTableName));
        if (string.IsNullOrWhiteSpace(toColumnName))
            throw new ArgumentException("toColumnName is required for create-relationship action", nameof(toColumnName));

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch => { commands.CreateRelationship(batch, fromTableName, fromColumnName, toTableName, toColumnName, isActive ?? true); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTableName}.{fromColumnName} to {toTableName}.{toColumnName} created successfully"
            }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string UpdateRelationshipAction(DataModelCommands commands, string sessionId, string? fromTableName, string? fromColumnName, string? toTableName, string? toColumnName, bool? isActive)
    {
        if (string.IsNullOrWhiteSpace(fromTableName))
            throw new ArgumentException("fromTableName is required for update-relationship action", nameof(fromTableName));
        if (string.IsNullOrWhiteSpace(fromColumnName))
            throw new ArgumentException("fromColumnName is required for update-relationship action", nameof(fromColumnName));
        if (string.IsNullOrWhiteSpace(toTableName))
            throw new ArgumentException("toTableName is required for update-relationship action", nameof(toTableName));
        if (string.IsNullOrWhiteSpace(toColumnName))
            throw new ArgumentException("toColumnName is required for update-relationship action", nameof(toColumnName));
        if (!isActive.HasValue)
            throw new ArgumentException("isActive is required for update-relationship action", nameof(isActive));

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch => { commands.UpdateRelationship(batch, fromTableName, fromColumnName, toTableName, toColumnName, isActive.Value); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTableName}.{fromColumnName} to {toTableName}.{toColumnName} updated successfully"
            }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string DeleteRelationshipAction(DataModelCommands commands, string sessionId, string? fromTableName, string? fromColumnName, string? toTableName, string? toColumnName)
    {
        if (string.IsNullOrWhiteSpace(fromTableName))
            throw new ArgumentException("fromTableName is required for delete-relationship action", nameof(fromTableName));
        if (string.IsNullOrWhiteSpace(fromColumnName))
            throw new ArgumentException("fromColumnName is required for delete-relationship action", nameof(fromColumnName));
        if (string.IsNullOrWhiteSpace(toTableName))
            throw new ArgumentException("toTableName is required for delete-relationship action", nameof(toTableName));
        if (string.IsNullOrWhiteSpace(toColumnName))
            throw new ArgumentException("toColumnName is required for delete-relationship action", nameof(toColumnName));

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch => { commands.DeleteRelationship(batch, fromTableName, fromColumnName, toTableName, toColumnName); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTableName}.{fromColumnName} to {toTableName}.{toColumnName} deleted successfully"
            }, ExcelToolsBase.JsonOptions);
        }
#pragma warning disable CA1031 // MCP protocol requires JSON error responses, not thrown exceptions
        catch (Exception ex)
#pragma warning restore CA1031
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }
}
