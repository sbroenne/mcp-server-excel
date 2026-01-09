using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Data Model relationship operations
/// </summary>
[McpServerToolType]
public static partial class ExcelDataModelRelTool
{
    /// <summary>
    /// Data Model relationship management.
    /// Related: excel_datamodel (tables/measures)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="sid">Session ID</param>
    /// <param name="ft">From table (source/many side)</param>
    /// <param name="fc">From column</param>
    /// <param name="tt">To table (target/one side)</param>
    /// <param name="tc">To column</param>
    /// <param name="active">Is relationship active (default: true)</param>
    [McpServerTool(Name = "excel_datamodel_rel", Title = "Excel Data Model Relationship Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelDataModelRel(
        DataModelRelAction action,
        string sid,
        [DefaultValue(null)] string? ft,
        [DefaultValue(null)] string? fc,
        [DefaultValue(null)] string? tt,
        [DefaultValue(null)] string? tc,
        [DefaultValue(null)] bool? active)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_datamodel_rel",
            action.ToActionString(),
            () =>
            {
                var commands = new DataModelCommands();

                return action switch
                {
                    DataModelRelAction.ListRelationships => ListRelationships(commands, sid),
                    DataModelRelAction.ReadRelationship => ReadRelationship(commands, sid, ft, fc, tt, tc),
                    DataModelRelAction.CreateRelationship => CreateRelationship(commands, sid, ft, fc, tt, tc, active),
                    DataModelRelAction.UpdateRelationship => UpdateRelationship(commands, sid, ft, fc, tt, tc, active),
                    DataModelRelAction.DeleteRelationship => DeleteRelationship(commands, sid, ft, fc, tt, tc),
                    _ => throw new ArgumentException($"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListRelationships(DataModelCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.ListRelationships(batch));

        return JsonSerializer.Serialize(new { result.Success, result.Relationships, result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadRelationship(DataModelCommands commands, string sessionId, string? fromTable, string? fromColumn, string? toTable, string? toColumn)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
            throw new ArgumentException("ft is required for read-relationship action", nameof(fromTable));
        if (string.IsNullOrWhiteSpace(fromColumn))
            throw new ArgumentException("fc is required for read-relationship action", nameof(fromColumn));
        if (string.IsNullOrWhiteSpace(toTable))
            throw new ArgumentException("tt is required for read-relationship action", nameof(toTable));
        if (string.IsNullOrWhiteSpace(toColumn))
            throw new ArgumentException("tc is required for read-relationship action", nameof(toColumn));

        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.ReadRelationship(batch, fromTable, fromColumn, toTable, toColumn));

        return JsonSerializer.Serialize(new { result.Success, result.Relationship, result.ErrorMessage }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateRelationship(DataModelCommands commands, string sessionId, string? fromTable, string? fromColumn, string? toTable, string? toColumn, bool? isActive)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
            throw new ArgumentException("ft is required for create-relationship action", nameof(fromTable));
        if (string.IsNullOrWhiteSpace(fromColumn))
            throw new ArgumentException("fc is required for create-relationship action", nameof(fromColumn));
        if (string.IsNullOrWhiteSpace(toTable))
            throw new ArgumentException("tt is required for create-relationship action", nameof(toTable));
        if (string.IsNullOrWhiteSpace(toColumn))
            throw new ArgumentException("tc is required for create-relationship action", nameof(toColumn));

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch => { commands.CreateRelationship(batch, fromTable, fromColumn, toTable, toColumn, isActive ?? true); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} created successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string UpdateRelationship(DataModelCommands commands, string sessionId, string? fromTable, string? fromColumn, string? toTable, string? toColumn, bool? isActive)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
            throw new ArgumentException("ft is required for update-relationship action", nameof(fromTable));
        if (string.IsNullOrWhiteSpace(fromColumn))
            throw new ArgumentException("fc is required for update-relationship action", nameof(fromColumn));
        if (string.IsNullOrWhiteSpace(toTable))
            throw new ArgumentException("tt is required for update-relationship action", nameof(toTable));
        if (string.IsNullOrWhiteSpace(toColumn))
            throw new ArgumentException("tc is required for update-relationship action", nameof(toColumn));
        if (!isActive.HasValue)
            throw new ArgumentException("active is required for update-relationship action", nameof(isActive));

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch => { commands.UpdateRelationship(batch, fromTable, fromColumn, toTable, toColumn, isActive.Value); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} updated successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string DeleteRelationship(DataModelCommands commands, string sessionId, string? fromTable, string? fromColumn, string? toTable, string? toColumn)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
            throw new ArgumentException("ft is required for delete-relationship action", nameof(fromTable));
        if (string.IsNullOrWhiteSpace(fromColumn))
            throw new ArgumentException("fc is required for delete-relationship action", nameof(fromColumn));
        if (string.IsNullOrWhiteSpace(toTable))
            throw new ArgumentException("tt is required for delete-relationship action", nameof(toTable));
        if (string.IsNullOrWhiteSpace(toColumn))
            throw new ArgumentException("tc is required for delete-relationship action", nameof(toColumn));

        try
        {
            ExcelToolsBase.WithSession(sessionId,
                batch => { commands.DeleteRelationship(batch, fromTable, fromColumn, toTable, toColumn); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} deleted successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, errorMessage = ex.Message, isError = true }, ExcelToolsBase.JsonOptions);
        }
    }
}
