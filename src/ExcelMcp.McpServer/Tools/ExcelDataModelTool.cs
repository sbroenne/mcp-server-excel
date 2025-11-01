using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;
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

            var actionString = action.ToActionString();

            return actionString switch
            {
                // Discovery operations
                "list-tables" => await ListTablesAsync(dataModelCommands, excelPath, batchId),
                "list-measures" => await ListMeasuresAsync(dataModelCommands, excelPath, batchId),
                "view-measure" => await ViewMeasureAsync(dataModelCommands, excelPath, measureName, batchId),
                "export-measure" => await ExportMeasureAsync(dataModelCommands, excelPath, measureName, outputPath, batchId),
                "list-relationships" => await ListRelationshipsAsync(dataModelCommands, excelPath, batchId),
                "refresh" => await RefreshAsync(dataModelCommands, excelPath, batchId),
                "delete-measure" => await DeleteMeasureAsync(dataModelCommands, excelPath, measureName, batchId),
                "delete-relationship" => await DeleteRelationshipAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, batchId),
                "view-table" => await ViewTableAsync(dataModelCommands, excelPath, tableName, batchId),
                "get-model-info" => await GetModelInfoAsync(dataModelCommands, excelPath, batchId),

                // DAX measures (requires Office 2016+)
                "create-measure" => await CreateMeasureComAsync(dataModelCommands, excelPath, tableName, measureName, daxFormula, formatString, description, batchId),
                "update-measure" => await UpdateMeasureComAsync(dataModelCommands, excelPath, measureName, daxFormula, formatString, description, batchId),

                // Relationships (requires Office 2016+)
                "create-relationship" => await CreateRelationshipComAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, isActive, batchId),
                "update-relationship" => await UpdateRelationshipComAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, isActive, batchId),

                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{actionString}'. Supported: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship, view-table, get-model-info, create-measure, update-measure, create-relationship, update-relationship")
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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Use Power Query to add tables to the Data Model"
            ];
            result.WorkflowHint = "List failed. Ensure file has Data Model and retry.";
            throw new ModelContextProtocol.McpException($"list-tables failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'list-measures' to see DAX measures",
            "Use 'list-relationships' to view table connections",
            "Use 'refresh' to update table data"
        ];
        result.WorkflowHint = "Tables listed. Next, explore measures or relationships.";

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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Verify the file contains tables loaded to Data Model",
                "Use 'list-tables' to check if Data Model has data",
                "Use Power Pivot to create measures if none exist"
            ];
            result.WorkflowHint = "List failed. Ensure Data Model contains tables and measures.";
            throw new ModelContextProtocol.McpException($"list-measures failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'view-measure' to see full DAX formulas",
            "Use 'export-measure' to save DAX to file",
            "Use 'list-tables' to see source tables"
        ];
        result.WorkflowHint = "Measures listed. Next, view or export DAX formulas.";

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
            async (batch) => await commands.ViewMeasureAsync(batch, measureName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Check that the measure name is correct",
                "Use 'list-measures' to see available measures",
                "Verify the Data Model contains tables with measures"
            ];
            result.WorkflowHint = "View failed. Ensure measure exists and retry.";
            throw new ModelContextProtocol.McpException($"view-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'export-measure' to save DAX to file",
            "Analyze DAX formula for optimization",
            "Use 'list-tables' to understand source tables"
        ];
        result.WorkflowHint = "Measure viewed. Next, export or analyze DAX formula.";

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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Check that the measure exists using 'list-measures'",
                "Verify the output path is writable",
                "Ensure the Data Model contains this measure"
            ];
            result.WorkflowHint = "Export failed. Ensure measure exists and path is valid.";
            throw new ModelContextProtocol.McpException($"export-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Review exported DAX formula",
            "Use exported DAX in other workbooks",
            "Version control the DAX file"
        ];
        result.WorkflowHint = "Measure exported. Next, review or reuse DAX formula.";

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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Verify the file contains tables loaded to Data Model",
                "Use 'list-tables' to check if multiple tables exist",
                "Use Power Pivot to create relationships if needed"
            ];
            result.WorkflowHint = "List failed. Ensure Data Model contains multiple tables with relationships.";
            throw new ModelContextProtocol.McpException($"list-relationships failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'list-tables' to see related tables",
            "Use 'list-measures' to see measures using relationships",
            "Verify relationship cardinality and filter direction"
        ];
        result.WorkflowHint = "Relationships listed. Next, explore tables or measures.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: true,
            async (batch) => await commands.RefreshAsync(batch));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions =
            [
                "Verify the file contains tables loaded to Data Model",
                "Use 'list-tables' to check what tables exist",
                "Check source connections and network connectivity"
            ];
            result.WorkflowHint = "Refresh failed. Ensure Data Model contains tables and connections are valid.";
            throw new ModelContextProtocol.McpException($"refresh failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions =
        [
            "Use 'list-tables' to verify record counts",
            "Use 'list-measures' to see updated calculations",
            "Validate Data Model integrity"
        ];
        result.WorkflowHint = "Data Model refreshed. Next, verify data and calculations.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Use 'list-measures' to see available measures",
                    "Check measure name for typos",
                    "Verify the file contains tables loaded to Data Model"
                ];
            }
            result.WorkflowHint = $"Failed to delete measure '{measureName}'. Verify measure exists.";
            throw new ModelContextProtocol.McpException($"delete-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                $"Measure '{measureName}' deleted successfully",
                "Use 'list-measures' to verify deletion",
                "Changes saved to workbook"
            ];
        }
        result.WorkflowHint = "Measure deleted. Next, verify remaining measures or create new ones.";

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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Use 'list-relationships' to see available relationships",
                    "Check table and column names for typos",
                    "Verify the file contains tables loaded to Data Model"
                ];
            }
            result.WorkflowHint = $"Failed to delete relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Verify relationship exists.";
            throw new ModelContextProtocol.McpException($"delete-relationship failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} deleted successfully",
                "Use 'list-relationships' to verify deletion",
                "Changes saved to workbook"
            ];
        }
        result.WorkflowHint = "Relationship deleted. Next, verify remaining relationships or create new ones.";

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
            async (batch) => await commands.ListTableColumnsAsync(batch, tableName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Use 'list-tables' to see available tables",
                    "Check table name for typos",
                    "Verify the file contains tables loaded to Data Model"
                ];
            }
            result.WorkflowHint = $"Failed to list columns for table '{tableName}'. Verify table exists.";
            throw new ModelContextProtocol.McpException($"list-columns failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                "Use 'view-table' to see table details",
                "Use 'create-measure' to create calculated measures",
                "Use 'create-relationship' to link tables"
            ];
        }
        result.WorkflowHint = $"Found columns in '{tableName}'. Use these for creating measures or relationships.";

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
            async (batch) => await commands.ViewTableAsync(batch, tableName));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Use 'list-tables' to see available tables",
                    "Check table name for typos",
                    "Verify the file contains tables loaded to Data Model"
                ];
            }
            result.WorkflowHint = $"Failed to view table '{tableName}'. Verify table exists.";
            throw new ModelContextProtocol.McpException($"view-table failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                "Use 'list-columns' to see all columns in detail",
                "Use 'create-measure' to add calculated measures",
                "Use 'list-relationships' to see table connections"
            ];
        }
        result.WorkflowHint = $"Viewed table '{tableName}'. Use this information for creating measures or relationships.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetModelInfoAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetModelInfoAsync(batch));

        // If operation failed, throw exception with detailed error message
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Verify the file contains tables loaded to Data Model",
                    "Try 'list-tables' to check if tables exist",
                    "Check if file is corrupted"
                ];
            }
            result.WorkflowHint = "Failed to get Data Model info. Verify workbook has a Data Model.";
            throw new ModelContextProtocol.McpException($"get-model-info failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                "Use 'list-tables' to explore tables",
                "Use 'list-measures' to see calculated measures",
                "Use 'list-relationships' to understand table connections"
            ];
        }
        result.WorkflowHint = "Got Data Model overview. Use list commands to explore in detail.";

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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Verify DAX formula syntax is correct",
                    "Use 'list-tables' to check available tables",
                    "Use 'list-columns' to see available columns",
                    "Check if measure name already exists"
                ];
            }
            result.WorkflowHint = $"Failed to create measure '{measureName}'. Check DAX syntax and table existence.";
            throw new ModelContextProtocol.McpException($"create-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                $"Measure '{measureName}' created successfully in table '{tableName}'",
                "Use 'view-measure' to verify the formula",
                "Use 'list-measures' to see all measures",
                "Changes saved to workbook"
            ];
        }
        result.WorkflowHint = "Measure created. Next, test it in a PivotTable or create more measures.";

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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Verify measure name exists using 'list-measures'",
                    "Check DAX formula syntax if updating formula",
                    "Use 'view-measure' to see current formula",
                    "Ensure at least one property is provided for update"
                ];
            }
            result.WorkflowHint = $"Failed to update measure '{measureName}'. Verify measure exists and DAX syntax.";
            throw new ModelContextProtocol.McpException($"update-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                $"Measure '{measureName}' updated successfully",
                "Use 'view-measure' to verify the changes",
                "Use 'list-measures' to see all measures",
                "Changes saved to workbook"
            ];
        }
        result.WorkflowHint = "Measure updated. Next, test the changes in a PivotTable.";

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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Use 'list-tables' to verify both tables exist",
                    "Use 'list-columns' to verify columns exist in both tables",
                    "Check if relationship already exists",
                    "Verify column data types are compatible"
                ];
            }
            result.WorkflowHint = $"Failed to create relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Verify tables and columns.";
            throw new ModelContextProtocol.McpException($"create-relationship failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                $"Relationship created from {fromTable}.{fromColumn} to {toTable}.{toColumn}",
                "Use 'list-relationships' to verify the relationship",
                "Create measures that use this relationship",
                "Changes saved to workbook"
            ];
        }
        result.WorkflowHint = "Relationship created. Next, create measures that leverage this relationship.";

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
        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions =
                [
                    "Use 'list-relationships' to verify relationship exists",
                    "Check table and column names for typos",
                    "Verify the file contains tables loaded to Data Model"
                ];
            }
            result.WorkflowHint = $"Failed to update relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Verify relationship exists.";
            throw new ModelContextProtocol.McpException($"update-relationship failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions =
            [
                $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} updated successfully",
                "Use 'list-relationships' to verify the changes",
                "Changes saved to workbook"
            ];
        }
        result.WorkflowHint = "Relationship updated. Changes will affect measures using this relationship.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
