using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// MCP tool for Excel Data Model (Power Pivot) - tables and DAX measures.
/// Use excel_datamodel_rel for relationships.
/// </summary>
[McpServerToolType]
public static partial class ExcelDataModelTool
{
    /// <summary>
    /// Data Model (Power Pivot) - DAX measures and table management.
    ///
    /// PREREQUISITE: Tables must be added to the Data Model first.
    /// Use excel_table with add-to-datamodel action to add worksheet tables,
    /// or excel_powerquery to import and load data directly to the Data Model.
    ///
    /// DAX MEASURES:
    /// - Create measures with DAX formulas like 'SUM(Sales[Amount])'
    /// - Measures can reference columns, other measures, and use DAX functions
    /// - Format string uses US format codes like '#,##0.00' for currency
    /// - DAX formulas are automatically formatted with proper indentation (daxformatter.com)
    ///
    /// DESTRUCTIVE OPERATIONS:
    /// - delete-table: Removes table AND all its measures - cannot be undone
    ///
    /// TIMEOUT: Operations auto-timeout after 2 minutes for large Data Models.
    ///
    /// RELATED TOOLS:
    /// - excel_table: Add worksheet tables to Data Model (add-to-datamodel action)
    /// - excel_datamodel_rel: Manage relationships between tables
    /// </summary>
    /// <param name="action">The Data Model operation to perform</param>
    /// <param name="sessionId">Session identifier returned from excel_file open action</param>
    /// <param name="measureName">Name of the DAX measure - required for measure operations</param>
    /// <param name="tableName">Name of the table in the Data Model - required for table operations and create-measure</param>
    /// <param name="newTableName">New name for rename-table action</param>
    /// <param name="daxFormula">DAX formula for the measure, e.g., 'SUM(Sales[Amount])'</param>
    /// <param name="description">Optional description for the measure</param>
    /// <param name="formatString">Number format code in US format, e.g., '#,##0.00' for currency</param>
    [McpServerTool(Name = "excel_datamodel", Title = "Excel Data Model Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelDataModel(
        DataModelAction action,
        string sessionId,
        [DefaultValue(null)] string? measureName,
        [DefaultValue(null)] string? tableName,
        [DefaultValue(null)] string? newTableName,
        [DefaultValue(null)] string? daxFormula,
        [DefaultValue(null)] string? description,
        [DefaultValue(null)] string? formatString)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_datamodel",
            action.ToActionString(),
            () =>
            {
                var dataModelCommands = new DataModelCommands();

                return action switch
                {
                    DataModelAction.ListTables => ListTablesAction(dataModelCommands, sessionId),
                    DataModelAction.ListMeasures => ListMeasuresAction(dataModelCommands, sessionId),
                    DataModelAction.Read => ReadMeasureAction(dataModelCommands, sessionId, measureName),
                    DataModelAction.Refresh => RefreshAction(dataModelCommands, sessionId),
                    DataModelAction.DeleteMeasure => DeleteMeasureAction(dataModelCommands, sessionId, measureName),
                    DataModelAction.DeleteTable => DeleteTableAction(dataModelCommands, sessionId, tableName),
                    DataModelAction.RenameTable => RenameTableAction(dataModelCommands, sessionId, tableName, newTableName),
                    DataModelAction.ReadTable => ReadTableAction(dataModelCommands, sessionId, tableName),
                    DataModelAction.ListColumns => ListColumnsAction(dataModelCommands, sessionId, tableName),
                    DataModelAction.ReadInfo => ReadModelInfoAction(dataModelCommands, sessionId),
                    DataModelAction.CreateMeasure => CreateMeasureAction(dataModelCommands, sessionId, tableName, measureName, daxFormula, formatString, description),
                    DataModelAction.UpdateMeasure => UpdateMeasureAction(dataModelCommands, sessionId, measureName, daxFormula, formatString, description),
                    _ => throw new ArgumentException(
                        $"Unknown action: {action} ({action.ToActionString()})", nameof(action))
                };
            });
    }

    private static string ListTablesAction(DataModelCommands commands, string sessionId)
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

    private static string ListMeasuresAction(DataModelCommands commands, string sessionId)
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

    private static string ReadMeasureAction(DataModelCommands commands, string sessionId, string? measureName)
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

    private static string RefreshAction(DataModelCommands commands, string sessionId)
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

    private static string DeleteMeasureAction(DataModelCommands commands, string sessionId, string? measureName)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ArgumentException("measureName is required for delete-measure action", nameof(measureName));
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

    private static string DeleteTableAction(DataModelCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("tableName is required for delete-table action", nameof(tableName));
        }

        try
        {
            ExcelToolsBase.WithSession(
                sessionId,
                batch => { commands.DeleteTable(batch, tableName); return 0; });

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Table '{tableName}' deleted from Data Model successfully"
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

    private static string RenameTableAction(DataModelCommands commands, string sessionId, string? tableName, string? newTableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("tableName is required for rename-table action", nameof(tableName));
        }

        if (string.IsNullOrWhiteSpace(newTableName))
        {
            throw new ArgumentException("newTableName is required for rename-table action", nameof(newTableName));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.RenameTable(batch, tableName, newTableName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.OldName,
            result.NormalizedNewName,
            result.ErrorMessage,
            isError = !result.Success
        }, ExcelToolsBase.JsonOptions);
    }

    private static string ReadTableAction(DataModelCommands commands, string sessionId,
        string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("tableName is required for read-table action", nameof(tableName));
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

    private static string ListColumnsAction(DataModelCommands commands, string sessionId,
        string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("tableName is required for list-columns action", nameof(tableName));
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

    private static string ReadModelInfoAction(DataModelCommands commands, string sessionId)
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

    private static string CreateMeasureAction(DataModelCommands commands,
        string sessionId, string? tableName, string? measureName, string? daxFormula, string? formatString,
        string? description)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("tableName is required for create-measure action", nameof(tableName));
        }

        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ArgumentException("measureName is required for create-measure action", nameof(measureName));
        }

        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            throw new ArgumentException("daxFormula is required for create-measure action", nameof(daxFormula));
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

    private static string UpdateMeasureAction(DataModelCommands commands,
        string sessionId, string? measureName, string? daxFormula, string? formatString, string? description)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ArgumentException("measureName is required for update-measure action", nameof(measureName));
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

}
