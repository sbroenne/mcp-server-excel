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

⚠️ CALCULATED COLUMNS: NOT supported via automation. When user asks to create calculated columns:
  - Provide step-by-step manual instructions (see LLM Usage Patterns in code comments)
  - OR suggest using DAX measures instead (measures ARE automated and usually better for aggregations)
")]
    public static async Task<string> ExcelDataModel(
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
        [Description("Measure name (for view-measure, export-measure, delete-measure, update-measure)")]
        string? measureName = null,

        [FileExtensions(Extensions = "dax")]
        [Description("Output file path for DAX export (for export-measure)")]
        string? outputPath = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Table name (for create-measure, view-table)")]
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
        bool? isActive = null,

        [Description("Timeout in minutes for data model operations. Default: 2 minutes")]
        double? timeout = null)
    {
        try
        {
            var dataModelCommands = new DataModelCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                // Discovery operations
                DataModelAction.ListTables => await ListTablesAsync(dataModelCommands, sessionId),
                DataModelAction.ListMeasures => await ListMeasuresAsync(dataModelCommands, sessionId),
                DataModelAction.Get => await ViewMeasureAsync(dataModelCommands, sessionId, measureName),
                DataModelAction.ExportMeasure => await ExportMeasureAsync(dataModelCommands, sessionId, measureName, outputPath),
                DataModelAction.ListRelationships => await ListRelationshipsAsync(dataModelCommands, sessionId),
                DataModelAction.Refresh => await RefreshAsync(dataModelCommands, excelPath, timeout, sessionId),
                DataModelAction.DeleteMeasure => await DeleteMeasureAsync(dataModelCommands, sessionId, measureName),
                DataModelAction.DeleteRelationship => await DeleteRelationshipAsync(dataModelCommands, sessionId, fromTable, fromColumn, toTable, toColumn),
                DataModelAction.GetTable => await ViewTableAsync(dataModelCommands, sessionId, tableName),
                DataModelAction.ListColumns => await ListColumnsAsync(dataModelCommands, sessionId, tableName),
                DataModelAction.GetInfo => await GetModelInfoAsync(dataModelCommands, sessionId),

                // DAX measures (requires Office 2016+)
                DataModelAction.CreateMeasure => await CreateMeasureComAsync(dataModelCommands, sessionId, tableName, measureName, daxFormula, formatString, description),
                DataModelAction.UpdateMeasure => await UpdateMeasureComAsync(dataModelCommands, sessionId, measureName, daxFormula, formatString, description),

                // Relationships (requires Office 2016+)
                DataModelAction.CreateRelationship => await CreateRelationshipComAsync(dataModelCommands, sessionId, fromTable, fromColumn, toTable, toColumn, isActive),
                DataModelAction.UpdateRelationship => await UpdateRelationshipComAsync(dataModelCommands, sessionId, fromTable, fromColumn, toTable, toColumn, isActive),

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

    private static async Task<string> ListTablesAsync(DataModelCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListTablesAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Tables,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListMeasuresAsync(DataModelCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListMeasuresAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Measures,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewMeasureAsync(DataModelCommands commands, string sessionId, string? measureName)
    {
        if (string.IsNullOrEmpty(measureName))
            throw new ModelContextProtocol.McpException("measureName is required for view-measure action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetAsync(batch, measureName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.MeasureName,
            result.DaxFormula,
            result.TableName,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ExportMeasureAsync(DataModelCommands commands, string sessionId, string? measureName, string? outputPath)
    {
        if (string.IsNullOrEmpty(measureName))
            throw new ModelContextProtocol.McpException("measureName is required for export-measure action");

        if (string.IsNullOrEmpty(outputPath))
            throw new ModelContextProtocol.McpException("outputPath is required for export-measure action");

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ExportMeasureAsync(batch, measureName, outputPath));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FilePath,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListRelationshipsAsync(DataModelCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListRelationshipsAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Relationships,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshAsync(DataModelCommands commands, string filePath, double? timeoutMinutes, string sessionId)
    {
        try
        {
            var timeoutSpan = timeoutMinutes.HasValue ? (TimeSpan?)TimeSpan.FromMinutes(timeoutMinutes.Value) : null;
            var result = await ExcelToolsBase.WithSessionAsync(
                sessionId,
                async batch => await commands.RefreshAsync(batch, null, timeoutSpan));

            return JsonSerializer.Serialize(new
            {
                result.Success,
                result.ErrorMessage
            }, ExcelToolsBase.JsonOptions);
        }
        catch (TimeoutException ex)
        {
            // Enrich timeout error with operation-specific guidance (MCP layer responsibility)
            var result = new OperationResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = filePath,
                Action = "refresh",

                OperationContext = new Dictionary<string, object>
                {
                    { "OperationType", "DataModel.Refresh" },
                    { "RefreshScope", "EntireModel" },
                    { "TimeoutReached", true },
                    { "UsedMaxTimeout", ex.Message.Contains("maximum timeout") }
                },

                IsRetryable = !ex.Message.Contains("maximum timeout"),

                RetryGuidance = ex.Message.Contains("maximum timeout")
                    ? "Maximum timeout (5 minutes) reached. Do not retry entire model refresh - try refreshing individual tables or check data source performance."
                    : "Retry acceptable if transient. For large models, consider table-by-table refresh strategy."
            };

            // MCP layer: Add workflow guidance for LLMs
            var response = new
            {
                result.Success,
                result.ErrorMessage,
                result.FilePath,
                result.Action,
                result.OperationContext,
                result.IsRetryable,
                result.RetryGuidance
            };

            return JsonSerializer.Serialize(response, ExcelToolsBase.JsonOptions);
        }
    }

    private static async Task<string> DeleteMeasureAsync(DataModelCommands commands, string sessionId, string? measureName)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'measureName' is required for delete-measure action");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteMeasureAsync(batch, measureName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteRelationshipAsync(DataModelCommands commands, string sessionId,
        string? fromTable, string? fromColumn, string? toTable, string? toColumn)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
        {
            throw new ModelContextProtocol.McpException("Parameter 'fromTable' is required for delete-relationship action");
        }

        if (string.IsNullOrWhiteSpace(fromColumn))
        {
            throw new ModelContextProtocol.McpException("Parameter 'fromColumn' is required for delete-relationship action");
        }

        if (string.IsNullOrWhiteSpace(toTable))
        {
            throw new ModelContextProtocol.McpException("Parameter 'toTable' is required for delete-relationship action");
        }

        if (string.IsNullOrWhiteSpace(toColumn))
        {
            throw new ModelContextProtocol.McpException("Parameter 'toColumn' is required for delete-relationship action");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.DeleteRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewTableAsync(DataModelCommands commands, string sessionId,
        string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for view-table action");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetTableAsync(batch, tableName));

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

    private static async Task<string> ListColumnsAsync(DataModelCommands commands, string sessionId,
        string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for list-columns action");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.ListColumnsAsync(batch, tableName));

        // Add workflow hints
        var columnCount = result.Columns?.Count ?? 0;
        var calculatedCount = result.Columns?.Count(c => c.IsCalculated) ?? 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            result.TableName,
            result.Columns
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetModelInfoAsync(DataModelCommands commands, string sessionId)
    {
        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.GetInfoAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableCount,
            result.MeasureCount,
            result.RelationshipCount,
            result.TotalRows,
            result.TableNames,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateMeasureComAsync(DataModelCommands commands,
        string sessionId, string? tableName, string? measureName, string? daxFormula, string? formatString,
        string? description)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for create-measure action");
        }

        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'measureName' is required for create-measure action");
        }

        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            throw new ModelContextProtocol.McpException("Parameter 'daxFormula' is required for create-measure action");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CreateMeasureAsync(batch, tableName, measureName, daxFormula,
                formatString, description));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateMeasureComAsync(DataModelCommands commands,
        string sessionId, string? measureName, string? daxFormula, string? formatString, string? description)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'measureName' is required for update-measure action");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.UpdateMeasureAsync(batch, measureName, daxFormula, formatString, description));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateRelationshipComAsync(DataModelCommands commands,
        string sessionId, string? fromTable, string? fromColumn, string? toTable, string? toColumn, bool? isActive)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
        {
            throw new ModelContextProtocol.McpException("Parameter 'fromTable' is required for create-relationship action");
        }

        if (string.IsNullOrWhiteSpace(fromColumn))
        {
            throw new ModelContextProtocol.McpException("Parameter 'fromColumn' is required for create-relationship action");
        }

        if (string.IsNullOrWhiteSpace(toTable))
        {
            throw new ModelContextProtocol.McpException("Parameter 'toTable' is required for create-relationship action");
        }

        if (string.IsNullOrWhiteSpace(toColumn))
        {
            throw new ModelContextProtocol.McpException("Parameter 'toColumn' is required for create-relationship action");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.CreateRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn,
                isActive ?? true));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateRelationshipComAsync(DataModelCommands commands,
        string sessionId, string? fromTable, string? fromColumn, string? toTable, string? toColumn, bool? isActive)
    {
        if (string.IsNullOrWhiteSpace(fromTable))
        {
            throw new ModelContextProtocol.McpException("Parameter 'fromTable' is required for update-relationship action");
        }

        if (string.IsNullOrWhiteSpace(fromColumn))
        {
            throw new ModelContextProtocol.McpException("Parameter 'fromColumn' is required for update-relationship action");
        }

        if (string.IsNullOrWhiteSpace(toTable))
        {
            throw new ModelContextProtocol.McpException("Parameter 'toTable' is required for update-relationship action");
        }

        if (string.IsNullOrWhiteSpace(toColumn))
        {
            throw new ModelContextProtocol.McpException("Parameter 'toColumn' is required for update-relationship action");
        }

        if (!isActive.HasValue)
        {
            throw new ModelContextProtocol.McpException("Parameter 'isActive' is required for update-relationship action");
        }

        var result = await ExcelToolsBase.WithSessionAsync(
            sessionId,
            async batch => await commands.UpdateRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn,
                isActive.Value));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage
        }, ExcelToolsBase.JsonOptions);
    }
}
