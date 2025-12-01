using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Data Model (Power Pivot) operations - DAX measures, relationships, and data refresh.
/// </summary>
[McpServerToolType]
public static partial class ExcelDataModelTool
{
    /// <summary>
    /// Manage Excel Power Pivot (Data Model) - DAX measures, relationships, analytical model.
    /// CALCULATED COLUMNS: NOT supported via automation. When user asks to create calculated columns, provide step-by-step manual instructions OR suggest using DAX measures instead (measures ARE automated and usually better for aggregations).
    /// TIMEOUT SAFEGUARD: Listing tables/measures/info auto-timeouts after 5 minutes to prevent Power Pivot hangs.
    /// </summary>
    /// <param name="action">Action to perform</param>
    /// <param name="excelPath">Excel file path (.xlsx or .xlsm)</param>
    /// <param name="sessionId">Session ID from excel_file 'open' action</param>
    /// <param name="measureName">Measure name (for read, export-measure, delete-measure, update-measure)</param>
    /// <param name="tableName">Table name (for create-measure, read-table)</param>
    /// <param name="daxFormula">DAX formula (for create-measure, update-measure)</param>
    /// <param name="description">Description (for create-measure, update-measure)</param>
    /// <param name="formatString">Format string (for create-measure, update-measure), e.g., '#,##0.00', '0.00%'</param>
    /// <param name="fromTable">Source table name (for delete-relationship, create-relationship, update-relationship)</param>
    /// <param name="fromColumn">Source column name (for delete-relationship, create-relationship, update-relationship)</param>
    /// <param name="toTable">Target table name (for delete-relationship, create-relationship, update-relationship)</param>
    /// <param name="toColumn">Target column name (for delete-relationship, create-relationship, update-relationship)</param>
    /// <param name="isActive">Whether relationship is active (for create-relationship, update-relationship), default: true</param>
    [McpServerTool(Name = "excel_datamodel")]
    [McpMeta("category", "analysis")]
    public static partial string ExcelDataModel(
        DataModelAction action,
        string excelPath,
        string sessionId,
        [DefaultValue(null)] string? measureName,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] string? daxFormula,
        [DefaultValue(null)] string? description,
        [DefaultValue(null)] string? formatString,
        [DefaultValue(null)] string? fromTable,
        [DefaultValue(null)] string? fromColumn,
        [DefaultValue(null)] string? toTable,
        [DefaultValue(null)] string? toColumn,
        [DefaultValue(null)] bool? isActive)
    {
        _ = excelPath; // retained for schema compatibility (operations require open session)

        return ExcelToolsBase.ExecuteToolAction(
            "excel_datamodel",
            action.ToActionString(),
            () =>
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
            });
    }

    private static string ListTablesAsync(DataModelCommands commands, string sessionId)
    {
        DataModelTableListResult result;

        try
        {
            result = ExcelToolsBase.WithSession(sessionId, batch => commands.ListTables(batch));
        }
        catch (TimeoutException ex)
        {
            result = new DataModelTableListResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                Tables = []
            };
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Tables,
            result.ErrorMessage,
            isError = !result.Success
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ListMeasuresAsync(DataModelCommands commands, string sessionId)
    {
        DataModelMeasureListResult result;

        try
        {
            result = ExcelToolsBase.WithSession(
                sessionId,
                batch => commands.ListMeasures(batch));
        }
        catch (TimeoutException ex)
        {
            result = new DataModelMeasureListResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                Measures = []
            };
        }

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Measures,
            result.ErrorMessage,
            isError = !result.Success
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
        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch => { commands.Refresh(batch, null, null); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Data Model refreshed successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string DeleteMeasureAsync(DataModelCommands commands, string sessionId, string? measureName)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ArgumentException("Parameter 'measureName' is required for delete-measure action", nameof(measureName));
        }

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch => { commands.DeleteMeasure(batch, measureName); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Measure '{measureName}' deleted successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch => { commands.DeleteRelationship(batch, fromTable, fromColumn, toTable, toColumn); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} deleted successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            result = ExcelToolsBase.WithSession(sessionId, batch => commands.ReadInfo(batch));
        }
        catch (TimeoutException ex)
        {
            result = new DataModelInfoResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                TableNames = []
            };
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
            isError = !result.Success
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch => { commands.CreateMeasure(batch, tableName, measureName, daxFormula, formatString, description); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Measure '{measureName}' created successfully in table '{tableName}'"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }

    private static string UpdateMeasureComAsync(DataModelCommands commands,
        string sessionId, string? measureName, string? daxFormula, string? formatString, string? description)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ArgumentException("Parameter 'measureName' is required for update-measure action", nameof(measureName));
        }

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch => { commands.UpdateMeasure(batch, measureName, daxFormula, formatString, description); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Measure '{measureName}' updated successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch => { commands.CreateRelationship(batch, fromTable, fromColumn, toTable, toColumn, isActive ?? true); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} created successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
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

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch => { commands.UpdateRelationship(batch, fromTable, fromColumn, toTable, toColumn, isActive.Value); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} updated successfully"
            }, ExcelToolsBase.JsonOptions);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                errorMessage = ex.Message,
                isError = true
            }, ExcelToolsBase.JsonOptions);
        }
    }
}

