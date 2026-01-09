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
    /// Data Model tables and DAX measures.
    /// DESTRUCTIVE: DeleteTable removes table AND all measures.
    /// TIMEOUT: 2 min auto-timeout.
    /// Related: excel_datamodel_rel (relationships)
    /// </summary>
    /// <param name="action">Action</param>
    /// <param name="sid">Session ID</param>
    /// <param name="mn">Measure name</param>
    /// <param name="tn">Table name</param>
    /// <param name="nn">New name (rename-table)</param>
    /// <param name="dax">DAX formula</param>
    /// <param name="desc">Description</param>
    /// <param name="fmt">Format string (#,##0.00)</param>
    [McpServerTool(Name = "excel_datamodel", Title = "Excel Data Model Operations")]
    [McpMeta("category", "analysis")]
    [McpMeta("requiresSession", true)]
    public static partial string ExcelDataModel(
        DataModelAction action,
        string sid,
        [DefaultValue(null)] string? mn,
        [DefaultValue(null)] string? tn,
        [DefaultValue(null)] string? nn,
        [DefaultValue(null)] string? dax,
        [DefaultValue(null)] string? desc,
        [DefaultValue(null)] string? fmt)
    {
        return ExcelToolsBase.ExecuteToolAction(
            "excel_datamodel",
            action.ToActionString(),
            () =>
            {
                var dataModelCommands = new DataModelCommands();

                return action switch
                {
                    DataModelAction.ListTables => ListTablesAsync(dataModelCommands, sid),
                    DataModelAction.ListMeasures => ListMeasuresAsync(dataModelCommands, sid),
                    DataModelAction.Read => ReadMeasureAsync(dataModelCommands, sid, mn),
                    DataModelAction.Refresh => RefreshAsync(dataModelCommands, sid),
                    DataModelAction.DeleteMeasure => DeleteMeasureAsync(dataModelCommands, sid, mn),
                    DataModelAction.DeleteTable => DeleteTableAsync(dataModelCommands, sid, tn),
                    DataModelAction.RenameTable => RenameTableAsync(dataModelCommands, sid, tn, nn),
                    DataModelAction.ReadTable => ReadTableAsync(dataModelCommands, sid, tn),
                    DataModelAction.ListColumns => ListColumnsAsync(dataModelCommands, sid, tn),
                    DataModelAction.ReadInfo => ReadModelInfoAsync(dataModelCommands, sid),
                    DataModelAction.CreateMeasure => CreateMeasureComAsync(dataModelCommands, sid, tn, mn, dax, fmt, desc),
                    DataModelAction.UpdateMeasure => UpdateMeasureComAsync(dataModelCommands, sid, mn, dax, fmt, desc),
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

    private static string DeleteTableAsync(DataModelCommands commands, string sessionId, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Parameter 'tableName' is required for delete-table action", nameof(tableName));
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

    private static string RenameTableAsync(DataModelCommands commands, string sessionId, string? tableName, string? newName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ArgumentException("Parameter 'tableName' is required for rename-table action", nameof(tableName));
        }

        if (string.IsNullOrWhiteSpace(newName))
        {
            throw new ArgumentException("Parameter 'newName' is required for rename-table action", nameof(newName));
        }

        var result = ExcelToolsBase.WithSession(
            sessionId,
            batch => commands.RenameTable(batch, tableName, newName));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.OldName,
            result.NormalizedNewName,
            result.ErrorMessage,
            isError = !result.Success
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

}
