using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.McpServer.Models;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel Data Model management tool for MCP server.
/// Provides access to Power Pivot Data Model operations.
///
/// LLM Usage Patterns:
///
/// DISCOVERY:
/// - Use "list-tables" to see all tables in the Data Model
/// - Use "list-measures" to view all DAX measures
/// - Use "list-relationships" to see table relationships
/// - Use "view-table" to see detailed table information
/// - Use "view-measure" to inspect DAX formula for a specific measure
/// - Use "get-model-info" to get Data Model overview
///
/// DAX MEASURES
/// - Use "create-measure" to add new DAX measures with optional format strings
/// - Use "update-measure" to modify existing measure formulas or formats
/// - Use "delete-measure" to remove a measure
/// - Use "export-measure" to save DAX formula to a file
///
/// RELATIONSHIPS
/// - Use "create-relationship" to define table relationships
/// - Use "update-relationship" to modify relationship active status
/// - Use "delete-relationship" to remove a relationship
///
/// DATA REFRESH:
/// - Use "refresh" to update Data Model data from source connections
///
/// CALCULATED COLUMNS (MANUAL ONLY):
/// - Calculated columns CANNOT be created via automation
/// - When user asks to create calculated columns, provide these EXACT instructions:
///
///   "To create a calculated column in Excel's Data Model:
///
///   1. Click on the Data Model table tab at the bottom of the Excel window
///   2. OR: Go to Power Pivot tab → Manage Data Model
///   3. In Power Pivot window, select the table (e.g., 'Sales')
///   4. Click in the 'Add Column' column header
///   5. Type your DAX formula (e.g., '=[Revenue] - [Cost]')
///   6. Press Enter
///   7. Right-click the column header → Rename Column
///   8. Set the column name (e.g., 'Profit')
///   9. Close Power Pivot window to save changes
///
///   The calculated column will now be available in PivotTables and DAX measures."
///
/// - Alternative approach: Guide user to create DAX measures instead (measures are automated)
/// - Measures are usually preferred over calculated columns for aggregations
/// </summary>
[McpServerToolType]
public static class ExcelDataModelTool
{
    /// <summary>
    /// Manage Excel Data Model (Power Pivot) - tables, measures, relationships
    /// </summary>
    [McpServerTool(Name = "excel_datamodel")]
    [Description(@"Manage Excel Power Pivot (Data Model) - DAX measures, relationships, analytical model.

⚡ PERFORMANCE: For creating 2+ measures/relationships, use begin_excel_batch FIRST (75-90% faster):
  1. batch = begin_excel_batch(excelPath: 'file.xlsx')
  2. excel_datamodel(action: 'create-measure', ..., batchId: batch.batchId)  // repeat
  3. commit_excel_batch(batchId: batch.batchId, save: true)

KEYWORDS: Power Pivot, PowerPivot, Data Model, DAX, measures, relationships, calculated columns.

⚠️ CALCULATED COLUMNS: NOT supported via automation. When user asks to create calculated columns:
  - Provide step-by-step manual instructions (see LLM Usage Patterns in code comments)
  - OR suggest using DAX measures instead (measures ARE automated and usually better for aggregations)

TYPICAL WORKFLOW:
1. Load data: excel_powerquery(action: 'set-load-to-data-model') ← loads to Power Pivot
2. Create relationships: excel_datamodel(action: 'create-relationship')
3. Create DAX measures: excel_datamodel(action: 'create-measure')

Actions: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship, view-table, get-model-info, create-measure, update-measure, create-relationship, update-relationship.")]
    public static async Task<string> ExcelDataModel(
        [Required]
        [Description("Action to perform (enum displayed as dropdown in MCP clients)")]
        DataModelAction action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

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

        [RegularExpression("^(Single|Both)$")]
        [Description("Cross-filter direction (for create-relationship, update-relationship): Single (default), Both")]
        string? crossFilterDirection = null,

        [Description("Optional batch ID for grouping operations")]
        string? batchId = null)
    {
        try
        {
            var dataModelCommands = new DataModelCommands();

            // Switch directly on enum for compile-time exhaustiveness checking (CS8524)
            return action switch
            {
                // Discovery operations
                DataModelAction.ListTables => await ListTablesAsync(dataModelCommands, excelPath, batchId),
                DataModelAction.ListMeasures => await ListMeasuresAsync(dataModelCommands, excelPath, batchId),
                DataModelAction.Get => await ViewMeasureAsync(dataModelCommands, excelPath, measureName, batchId),
                DataModelAction.ExportMeasure => await ExportMeasureAsync(dataModelCommands, excelPath, measureName, outputPath, batchId),
                DataModelAction.ListRelationships => await ListRelationshipsAsync(dataModelCommands, excelPath, batchId),
                DataModelAction.Refresh => await RefreshAsync(dataModelCommands, excelPath, batchId),
                DataModelAction.DeleteMeasure => await DeleteMeasureAsync(dataModelCommands, excelPath, measureName, batchId),
                DataModelAction.DeleteRelationship => await DeleteRelationshipAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, batchId),
                DataModelAction.GetTable => await ViewTableAsync(dataModelCommands, excelPath, tableName, batchId),
                DataModelAction.GetInfo => await GetModelInfoAsync(dataModelCommands, excelPath, batchId),

                // DAX measures (requires Office 2016+)
                DataModelAction.CreateMeasure => await CreateMeasureComAsync(dataModelCommands, excelPath, tableName, measureName, daxFormula, formatString, description, batchId),
                DataModelAction.UpdateMeasure => await UpdateMeasureComAsync(dataModelCommands, excelPath, measureName, daxFormula, formatString, description, batchId),

                // Relationships (requires Office 2016+)
                DataModelAction.CreateRelationship => await CreateRelationshipComAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, isActive, batchId),
                DataModelAction.UpdateRelationship => await UpdateRelationshipComAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, isActive, batchId),

                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action: {action} ({action.ToActionString()})")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action.ToActionString(), excelPath);
            throw; // Unreachable but satisfies compiler
        }
    }

    private static async Task<string> ListTablesAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListTablesAsync(batch));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListMeasuresAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListMeasuresAsync(batch));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewMeasureAsync(DataModelCommands commands, string filePath, string? measureName, string? batchId)
    {
        if (string.IsNullOrEmpty(measureName))
            throw new ModelContextProtocol.McpException("measureName is required for view-measure action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetAsync(batch, measureName));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ExportMeasureAsync(DataModelCommands commands, string filePath, string? measureName, string? outputPath, string? batchId)
    {
        if (string.IsNullOrEmpty(measureName))
            throw new ModelContextProtocol.McpException("measureName is required for export-measure action");

        if (string.IsNullOrEmpty(outputPath))
            throw new ModelContextProtocol.McpException("outputPath is required for export-measure action");

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ExportMeasureAsync(batch, measureName, outputPath));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListRelationshipsAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListRelationshipsAsync(batch));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        try
        {
            var result = await ExcelToolsBase.WithBatchAsync(
                batchId,
                filePath,
                save: true,
                async (batch) => await commands.RefreshAsync(batch));

            // If operation failed, throw exception with detailed error message
            // Always return JSON (success or failure) - MCP clients handle the success flag
            return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
        }
        catch (TimeoutException ex)
        {
            // Enrich timeout error with operation-specific guidance
            var result = new OperationResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = filePath,
                Action = "refresh",

                SuggestedNextActions = new List<string>
                {
                    "Check if Excel is showing a dialog or is unresponsive",
                    "Verify all data source connections in the Data Model are accessible",
                    "For large Data Models (millions of rows), refresh may genuinely require 5+ minutes",
                    "Consider refreshing individual tables instead of entire model (use tableName parameter)"
                },

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

            return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
        }
    }

    private static async Task<string> DeleteMeasureAsync(DataModelCommands commands, string filePath, string? measureName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'measureName' is required for delete-measure action");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteMeasureAsync(batch, measureName));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteRelationshipAsync(DataModelCommands commands, string filePath,
        string? fromTable, string? fromColumn, string? toTable, string? toColumn, string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.DeleteRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListTableColumnsAsync(DataModelCommands commands, string filePath,
        string? tableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for list-columns action");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListColumnsAsync(batch, tableName));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewTableAsync(DataModelCommands commands, string filePath,
        string? tableName, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for view-table action");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetTableAsync(batch, tableName));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetModelInfoAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetInfoAsync(batch));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateMeasureComAsync(DataModelCommands commands, string filePath,
        string? tableName, string? measureName, string? daxFormula, string? formatString,
        string? description, string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CreateMeasureAsync(batch, tableName, measureName, daxFormula,
                formatString, description));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateMeasureComAsync(DataModelCommands commands, string filePath,
        string? measureName, string? daxFormula, string? formatString, string? description, string? batchId)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'measureName' is required for update-measure action");
        }

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.UpdateMeasureAsync(batch, measureName, daxFormula, formatString, description));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateRelationshipComAsync(DataModelCommands commands, string filePath,
        string? fromTable, string? fromColumn, string? toTable, string? toColumn, bool? isActive, string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.CreateRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn,
                isActive ?? true));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateRelationshipComAsync(DataModelCommands commands, string filePath,
        string? fromTable, string? fromColumn, string? toTable, string? toColumn, bool? isActive, string? batchId)
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

        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.UpdateRelationshipAsync(batch, fromTable, fromColumn, toTable, toColumn,
                isActive.Value));

        // If operation failed, throw exception with detailed error message
        // Always return JSON (success or failure) - MCP clients handle the success flag
        // Success - add workflow guidance


        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
