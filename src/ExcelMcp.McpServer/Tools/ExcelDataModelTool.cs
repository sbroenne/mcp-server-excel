using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

#pragma warning disable CA1861 // Avoid constant arrays as arguments - workflow hints are contextual per-call

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Data Model (Power Pivot) operations - DAX measures, relationships, and data refresh.
/// </summary>
[McpServerToolType]
public static class ExcelDataModelTool
{
    /// <summary>
    /// Manage Excel Data Model (Power Pivot) - tables, measures, relationships
    /// </summary>
    [McpServerTool(Name = "excel_datamodel")]
    [Description(@"Manage Excel Power Pivot (Data Model) - DAX measures, relationships, analytical model.

**CALCULATED COLUMNS:** NOT supported via automation. When user asks to create calculated columns:
  - Provide step-by-step manual instructions (see LLM Usage Patterns in code comments)
  - OR suggest using DAX measures instead (measures ARE automated and usually better for aggregations)

    ⏱️ TIMEOUT SAFEGUARD:
    - Listing tables/measures/info auto-timeouts after 5 minutes to avoid Power Pivot hangs
    - On timeout the tool returns SuggestedNextActions instead of freezing the MCP session
")]
    public static string ExcelDataModel(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        DataModelAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [Required]
        [Description("Session ID from excel_file 'open' action")]
        string sessionId,

        [StringLength(255, MinimumLength = 1)]
        [Description("Measure name (for read, export-measure, delete-measure, update-measure)")]
        string? measureName = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Table name (for create-measure, read-table)")]
        string? tableName = null,

        [StringLength(8000, MinimumLength = 1)]
        [Description("DAX formula (for create-measure, update-measure)")]
        string? daxFormula = null,

        [StringLength(1000)]
        [Description("Description (for create-measure, update-measure)")]
        string? description = null,

        [StringLength(255)]
        [Description("Format string (for create-measure, update-measure), e.g., '#,##0.00', '0.00%'")]
        string? formatString = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Source table name (for delete-relationship, create-relationship, update-relationship)")]
        string? fromTable = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Source column name (for delete-relationship, create-relationship, update-relationship)")]
        string? fromColumn = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Target table name (for delete-relationship, create-relationship, update-relationship)")]
        string? toTable = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Target column name (for delete-relationship, create-relationship, update-relationship)")]
        string? toColumn = null,

        [Description("Whether relationship is active (for create-relationship, update-relationship), default: true")]
        bool? isActive = null)
    {
        _ = excelPath; // retained for schema compatibility (operations require open session)
        try
        {
            var dataModelCommands = new DataModelCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                // Discovery operations
                DataModelAction.ListTables => ListTablesAsync(dataModelCommands, sessionId),
                DataModelAction.ListMeasures => ListMeasuresAsync(dataModelCommands, sessionId),
                DataModelAction.Read => ReadMeasureAsync(dataModelCommands, sessionId, measureName),
                DataModelAction.ListRelationships => ListRelationshipsAsync(dataModelCommands, sessionId),
                DataModelAction.Refresh => RefreshAsync(dataModelCommands, sessionId),
                DataModelAction.DeleteMeasure => DeleteMeasureAsync(dataModelCommands, sessionId, measureName),
                DataModelAction.DeleteRelationship => DeleteRelationshipAsync(dataModelCommands, sessionId, fromTable, fromColumn, toTable, toColumn),
                DataModelAction.ReadTable => ReadTableAsync(dataModelCommands, sessionId, tableName),
                DataModelAction.ListColumns => ListColumnsAsync(dataModelCommands, sessionId, tableName),
                DataModelAction.ReadInfo => ReadModelInfoAsync(dataModelCommands, sessionId),

                // DAX measures (requires Office 2016+)
                DataModelAction.CreateMeasure => CreateMeasureComAsync(dataModelCommands, sessionId, tableName, measureName, daxFormula, formatString, description),
                DataModelAction.UpdateMeasure => UpdateMeasureComAsync(dataModelCommands, sessionId, measureName, daxFormula, formatString, description),

                // Relationships (requires Office 2016+)
                DataModelAction.CreateRelationship => CreateRelationshipComAsync(dataModelCommands, sessionId, fromTable, fromColumn, toTable, toColumn, isActive),
                DataModelAction.UpdateRelationship => UpdateRelationshipComAsync(dataModelCommands, sessionId, fromTable, fromColumn, toTable, toColumn, isActive),

                _ => throw new ArgumentException(
                    $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
            };
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = $"{action.ToActionString()} failed: {ex.Message}",
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string ListTablesAsync(DataModelCommands commands, string sessionId)
    {
        DataModelTableListResult result;
        bool isTimeout = false;
        string[]? suggestedNextActions = null;
        Dictionary<string, object>? operationContext = null;
        string? retryGuidance = null;

        try
        {
            result = ExcelToolsBase.WithSession(sessionId, batch => commands.ListTables(batch));
        }
        catch (TimeoutException ex)
        {
            isTimeout = true;
            bool usedMaxTimeout = ex.Message.Contains("maximum timeout", StringComparison.OrdinalIgnoreCase);

            result = new DataModelTableListResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                Tables = []
            };

            suggestedNextActions =
            [
                "Check Excel for Power Pivot dialogs requesting credentials or privacy confirmations.",
                "Refresh the Data Model manually to ensure tables are available.",
                "If the model is large, filter or remove unused tables before listing again.",
                "Use begin_excel_batch when enumerating multiple Data Model actions to keep the session alive."
            ];

            operationContext = new Dictionary<string, object>
            {
                ["OperationType"] = "DataModel.ListTables",
                ["TimeoutReached"] = true,
                ["UsedMaxTimeout"] = usedMaxTimeout
            };

            retryGuidance = usedMaxTimeout
                ? "Maximum timeout reached. Open Excel's Power Pivot window to resolve prompts before retrying."
                : "After verifying the model is responsive, retry listing tables (up to the 5 minute timeout).";
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Tables,
            result.ErrorMessage,
            isError = !result.Success || isTimeout,
            suggestedNextActions,
            retryGuidance,
            operationContext
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ListMeasuresAsync(DataModelCommands commands, string sessionId)
    {
        DataModelMeasureListResult result;
        bool isTimeout = false;
        string[]? suggestedNextActions = null;
        Dictionary<string, object>? operationContext = null;
        string? retryGuidance = null;

        try
        {
            result = ExcelToolsBase.WithSession(
                sessionId,
                batch => commands.ListMeasures(batch));
        }
        catch (TimeoutException ex)
        {
            isTimeout = true;
            bool usedMaxTimeout = ex.Message.Contains("maximum timeout", StringComparison.OrdinalIgnoreCase);

            result = new DataModelMeasureListResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                Measures = []
            };

            suggestedNextActions =
            [
                "Open Power Pivot to verify measures load without prompts.",
                "Refresh Data Model connections to ensure measures are accessible.",
                "Use excel_datamodel(action: 'list-tables') to confirm table metadata before listing measures.",
                "Keep the session warm with begin_excel_batch while inspecting multiple Data Model objects."
            ];

            operationContext = new Dictionary<string, object>
            {
                ["OperationType"] = "DataModel.ListMeasures",
                ["TimeoutReached"] = true,
                ["UsedMaxTimeout"] = usedMaxTimeout
            };

            retryGuidance = usedMaxTimeout
                ? "Maximum timeout reached. Resolve Power Pivot prompts or reduce model complexity before retrying."
                : "After confirming the model is responsive, retry listing measures within the 5 minute timeout.";
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Measures,
            result.ErrorMessage,
            isError = !result.Success || isTimeout,
            suggestedNextActions,
            retryGuidance,
            operationContext
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadMeasureAsync(DataModelCommands commands, string sessionId, string? measureName)
    {
        if (string.IsNullOrEmpty(measureName))
            throw new ArgumentException("measureName is required for read action", nameof(measureName));

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Read(batch, measureName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.MeasureName,
            result.DaxFormula,
            result.TableName,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ListRelationshipsAsync(DataModelCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(sessionId, batch => commands.ListRelationships(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Relationships,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string RefreshAsync(DataModelCommands commands, string sessionId)
    {
        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.Refresh(batch, null, null));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteMeasureAsync(DataModelCommands commands, string sessionId, string? measureName)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ArgumentException("Parameter 'measureName' is required for delete-measure action", nameof(measureName));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.DeleteMeasure(batch, measureName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteRelationshipAsync(DataModelCommands commands, string sessionId,
        string? fromTable, string? fromColumn, string? toTable, string? toColumn)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
        {
            throw new ArgumentException("Parameter 'fromTable' is required for delete-relationship action", nameof(fromTable));
        }

        if (string.IsNullOrWhiteSpace(fromColumn))
        {
            throw new ArgumentException("Parameter 'fromColumn' is required for delete-relationship action", nameof(fromColumn));
        }

        if (string.IsNullOrWhiteSpace(toTable))
        {
            throw new ArgumentException("Parameter 'toTable' is required for delete-relationship action", nameof(toTable));
        }

        if (string.IsNullOrWhiteSpace(toColumn))
        {
            throw new ArgumentException("Parameter 'toColumn' is required for delete-relationship action", nameof(toColumn));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.DeleteRelationship(batch, fromTable, fromColumn, toTable, toColumn));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadTableAsync(DataModelCommands commands, string sessionId,
        string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Parameter 'tableName' is required for read-table action", nameof(tableName));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ReadTable(batch, tableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableName,
            result.SourceName,
            result.RecordCount,
            result.Columns,
            result.MeasureCount,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ListColumnsAsync(DataModelCommands commands, string sessionId,
        string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Parameter 'tableName' is required for list-columns action", nameof(tableName));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.ListColumns(batch, tableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            result.TableName,
            result.Columns
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadModelInfoAsync(DataModelCommands commands, string sessionId)
    {
        DataModelInfoResult result;
        bool isTimeout = false;
        string[]? suggestedNextActions = null;
        Dictionary<string, object>? operationContext = null;
        string? retryGuidance = null;

        try
        {
            result = ExcelToolsBase.WithSession(sessionId, batch => commands.ReadInfo(batch));
        }
        catch (TimeoutException ex)
        {
            isTimeout = true;
            bool usedMaxTimeout = ex.Message.Contains("maximum timeout", StringComparison.OrdinalIgnoreCase);

            result = new DataModelInfoResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                TableNames = []
            };

            suggestedNextActions =
            [
                "Open Power Pivot to inspect the model and clear any pending dialogs.",
                "Refresh connections feeding the Data Model before requesting summary info.",
                "If the model is large, gather table-specific info instead of the full summary.",
                "Use begin_excel_batch to reuse the same Excel session while gathering metrics."
            ];

            operationContext = new Dictionary<string, object>
            {
                ["OperationType"] = "DataModel.ReadInfo",
                ["TimeoutReached"] = true,
                ["UsedMaxTimeout"] = usedMaxTimeout
            };

            retryGuidance = usedMaxTimeout
                ? "Maximum timeout reached. Resolve Power Pivot latency or dialogs before retrying."
                : "After confirming the model responds quickly, retry reading summary info within the timeout.";
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableCount,
            result.MeasureCount,
            result.RelationshipCount,
            result.TotalRows,
            result.TableNames,
            result.ErrorMessage,
            isError = !result.Success || isTimeout,
            suggestedNextActions,
            retryGuidance,
            operationContext
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateMeasureComAsync(DataModelCommands commands,
        string sessionId, string? tableName, string? measureName, string? daxFormula, string? formatString,
        string? description)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Parameter 'tableName' is required for create-measure action", nameof(tableName));
        }

        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ArgumentException("Parameter 'measureName' is required for create-measure action", nameof(measureName));
        }

        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            throw new ArgumentException("Parameter 'daxFormula' is required for create-measure action", nameof(daxFormula));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CreateMeasure(batch, tableName, measureName, daxFormula,
                formatString, description));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateMeasureComAsync(DataModelCommands commands,
        string sessionId, string? measureName, string? daxFormula, string? formatString, string? description)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ArgumentException("Parameter 'measureName' is required for update-measure action", nameof(measureName));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.UpdateMeasure(batch, measureName, daxFormula, formatString, description));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string CreateRelationshipComAsync(DataModelCommands commands,
        string sessionId, string? fromTable, string? fromColumn, string? toTable, string? toColumn, bool? isActive)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
        {
            throw new ArgumentException("Parameter 'fromTable' is required for create-relationship action", nameof(fromTable));
        }

        if (string.IsNullOrWhiteSpace(fromColumn))
        {
            throw new ArgumentException("Parameter 'fromColumn' is required for create-relationship action", nameof(fromColumn));
        }

        if (string.IsNullOrWhiteSpace(toTable))
        {
            throw new ArgumentException("Parameter 'toTable' is required for create-relationship action", nameof(toTable));
        }

        if (string.IsNullOrWhiteSpace(toColumn))
        {
            throw new ArgumentException("Parameter 'toColumn' is required for create-relationship action", nameof(toColumn));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.CreateRelationship(batch, fromTable, fromColumn, toTable, toColumn,
                isActive ?? true));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static string UpdateRelationshipComAsync(DataModelCommands commands,
        string sessionId, string? fromTable, string? fromColumn, string? toTable, string? toColumn, bool? isActive)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
        {
            throw new ArgumentException("Parameter 'fromTable' is required for update-relationship action", nameof(fromTable));
        }

        if (string.IsNullOrWhiteSpace(fromColumn))
        {
            throw new ArgumentException("Parameter 'fromColumn' is required for update-relationship action", nameof(fromColumn));
        }

        if (string.IsNullOrWhiteSpace(toTable))
        {
            throw new ArgumentException("Parameter 'toTable' is required for update-relationship action", nameof(toTable));
        }

        if (string.IsNullOrWhiteSpace(toColumn))
        {
            throw new ArgumentException("Parameter 'toColumn' is required for update-relationship action", nameof(toColumn));
        }

        if (!isActive.HasValue)
        {
            throw new ArgumentException("Parameter 'isActive' is required for update-relationship action", nameof(isActive));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.UpdateRelationship(batch, fromTable, fromColumn, toTable, toColumn,
                isActive.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}

