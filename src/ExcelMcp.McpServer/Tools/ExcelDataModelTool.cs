using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel Data Model management tool for MCP server.
/// Provides access to Power Pivot Data Model operations.
///
/// LLM Usage Patterns:
/// - Use "list-tables" to see all tables in the Data Model
/// - Use "list-measures" to view all DAX measures
/// - Use "view-measure" to inspect DAX formula for a specific measure
/// - Use "export-measure" to save DAX formula to a file
/// - Use "list-relationships" to see table relationships
/// - Use "refresh" to update Data Model data
/// - Use "create-measure" to add new DAX measures (TOM API)
/// - Use "update-measure" to modify existing measures (TOM API)
/// - Use "create-relationship" to define table relationships (TOM API)
/// - Use "update-relationship" to modify relationships (TOM API)
/// - Use "create-column" to add calculated columns (TOM API)
/// - Use "list-columns" to view all calculated columns (TOM API)
/// - Use "view-column" to see calculated column details (TOM API)
/// - Use "update-column" to modify calculated columns (TOM API)
/// - Use "delete-column" to remove calculated columns (TOM API)
/// - Use "validate-dax" to check DAX syntax (TOM API)
///
/// Phase 1 Scope (COM API):
/// - Tables: List and view metadata
/// - Measures: List, view, export, and delete DAX formulas
/// - Relationships: List and delete relationships
/// - Refresh: Refresh Data Model connections
///
/// Phase 4 Scope (TOM API):
/// - Measures: Create and update with full DAX support
/// - Relationships: Create and update with advanced options
/// - Calculated Columns: Create DAX-based calculated columns
/// - DAX Validation: Syntax checking before creation
/// </summary>
[McpServerToolType]
public static class ExcelDataModelTool
{
    /// <summary>
    /// Manage Excel Data Model (Power Pivot) - tables, measures, relationships
    /// </summary>
    [McpServerTool(Name = "excel_datamodel")]
    [Description("Manage Excel Data Model operations. Phase 2 (COM API): list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship, list-columns, view-table, get-model-info, create-measure, update-measure, create-relationship, update-relationship. Phase 4 (TOM API): create-column, view-column, update-column, delete-column, validate-dax.")]
    public static async Task<string> ExcelDataModel(
        [Required]
        [RegularExpression("^(list-tables|list-measures|view-measure|export-measure|list-relationships|refresh|delete-measure|delete-relationship|list-columns|view-table|get-model-info|create-measure|update-measure|create-relationship|update-relationship|create-column|view-column|update-column|delete-column|validate-dax)$")]
        [Description("Action: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship, list-columns, view-table, get-model-info, create-measure, update-measure, create-relationship, update-relationship, create-column, view-column, update-column, delete-column, validate-dax")]
        string action,

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
        [Description("Table name (for create-measure, create-column, list-columns, view-column, update-column, delete-column)")]
        string? tableName = null,

        [StringLength(8000, MinimumLength = 1)]
        [Description("DAX formula (for create-measure, update-measure, create-column, update-column, validate-dax)")]
        string? daxFormula = null,

        [StringLength(1000)]
        [Description("Description (for create-measure, update-measure, create-column, update-column)")]
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

        [StringLength(255, MinimumLength = 1)]
        [Description("Column name (for create-column, view-column, update-column, delete-column)")]
        string? columnName = null,

        [RegularExpression("^(String|Integer|Double|Boolean|DateTime)$")]
        [Description("Data type (for create-column, update-column): String, Integer, Double, Boolean, DateTime")]
        string? dataType = null,
        
        [Description("Optional batch ID for grouping operations")]
        string? batchId = null)
    {
        try
        {
            var dataModelCommands = new DataModelCommands();
            var tomCommands = new DataModelTomCommands();

            return action.ToLowerInvariant() switch
            {
                // COM API operations (Phase 1 + Phase 2)
                "list-tables" => await ListTablesAsync(dataModelCommands, excelPath, batchId),
                "list-measures" => await ListMeasuresAsync(dataModelCommands, excelPath, batchId),
                "view-measure" => await ViewMeasureAsync(dataModelCommands, excelPath, measureName, batchId),
                "export-measure" => await ExportMeasureAsync(dataModelCommands, excelPath, measureName, outputPath, batchId),
                "list-relationships" => await ListRelationshipsAsync(dataModelCommands, excelPath, batchId),
                "refresh" => await RefreshAsync(dataModelCommands, excelPath, batchId),
                "delete-measure" => await DeleteMeasureAsync(dataModelCommands, excelPath, measureName, batchId),
                "delete-relationship" => await DeleteRelationshipAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, batchId),
                
                // Phase 2: Discovery operations (COM API)
                "list-columns" => await ListTableColumnsAsync(dataModelCommands, excelPath, tableName, batchId),
                "view-table" => await ViewTableAsync(dataModelCommands, excelPath, tableName, batchId),
                "get-model-info" => await GetModelInfoAsync(dataModelCommands, excelPath, batchId),
                
                // Phase 2: CREATE/UPDATE operations (COM API - Office 2016+)
                "create-measure" => await CreateMeasureComAsync(dataModelCommands, excelPath, tableName, measureName, daxFormula, formatString, description, batchId),
                "update-measure" => await UpdateMeasureComAsync(dataModelCommands, excelPath, measureName, daxFormula, formatString, description, batchId),
                "create-relationship" => await CreateRelationshipComAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, isActive, batchId),
                "update-relationship" => await UpdateRelationshipComAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, isActive, batchId),

                // TOM API operations (Phase 4 - Future)
                "create-column" => await CreateCalculatedColumnAsync(tomCommands, excelPath, tableName, columnName, daxFormula, description, dataType, batchId),
                "view-column" => await ViewCalculatedColumnAsync(tomCommands, excelPath, tableName, columnName, batchId),
                "update-column" => await UpdateCalculatedColumnAsync(tomCommands, excelPath, tableName, columnName, daxFormula, description, dataType, batchId),
                "delete-column" => await DeleteCalculatedColumnAsync(tomCommands, excelPath, tableName, columnName, batchId),
                "validate-dax" => await ValidateDaxAsync(tomCommands, excelPath, daxFormula, batchId),

                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship, list-columns, view-table, get-model-info, create-measure, update-measure, create-relationship, update-relationship, create-column, view-column, update-column, delete-column, validate-dax")
            };
        }
        catch (ModelContextProtocol.McpException)
        {
            throw; // Re-throw MCP exceptions as-is
        }
        catch (Exception ex)
        {
            ExcelToolsBase.ThrowInternalError(ex, action, excelPath);
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
            result.SuggestedNextActions = new List<string>
            {
                "Check that the Excel file exists and is accessible",
                "Verify the file contains a Data Model (Power Pivot)",
                "Use Power Query to add tables to the Data Model"
            };
            result.WorkflowHint = "List failed. Ensure file has Data Model and retry.";
            throw new ModelContextProtocol.McpException($"list-tables failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list-measures' to see DAX measures",
            "Use 'list-relationships' to view table connections",
            "Use 'refresh' to update table data"
        };
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
            result.SuggestedNextActions = new List<string>
            {
                "Check that the Excel file contains a Data Model",
                "Verify measures exist in the Data Model",
                "Use Power Pivot to create measures if none exist"
            };
            result.WorkflowHint = "List failed. Ensure Data Model has measures and retry.";
            throw new ModelContextProtocol.McpException($"list-measures failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'view-measure' to see full DAX formulas",
            "Use 'export-measure' to save DAX to file",
            "Use 'list-tables' to see source tables"
        };
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
            result.SuggestedNextActions = new List<string>
            {
                "Check that the measure name is correct",
                "Use 'list-measures' to see available measures",
                "Verify the Data Model contains this measure"
            };
            result.WorkflowHint = "View failed. Ensure measure exists and retry.";
            throw new ModelContextProtocol.McpException($"view-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'export-measure' to save DAX to file",
            "Analyze DAX formula for optimization",
            "Use 'list-tables' to understand source tables"
        };
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
            result.SuggestedNextActions = new List<string>
            {
                "Check that the measure exists using 'list-measures'",
                "Verify the output path is writable",
                "Ensure the Data Model contains this measure"
            };
            result.WorkflowHint = "Export failed. Ensure measure exists and path is valid.";
            throw new ModelContextProtocol.McpException($"export-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Review exported DAX formula",
            "Use exported DAX in other workbooks",
            "Version control the DAX file"
        };
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
            result.SuggestedNextActions = new List<string>
            {
                "Check that the Excel file contains a Data Model",
                "Verify relationships exist between tables",
                "Use Power Pivot to create relationships if needed"
            };
            result.WorkflowHint = "List failed. Ensure Data Model has relationships and retry.";
            throw new ModelContextProtocol.McpException($"list-relationships failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list-tables' to see related tables",
            "Use 'list-measures' to see measures using relationships",
            "Verify relationship cardinality and filter direction"
        };
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
            result.SuggestedNextActions = new List<string>
            {
                "Check that the Excel file contains a Data Model",
                "Verify source connections are valid",
                "Ensure network connectivity to data sources"
            };
            result.WorkflowHint = "Refresh failed. Ensure connections are valid and retry.";
            throw new ModelContextProtocol.McpException($"refresh failed for '{filePath}': {result.ErrorMessage}");
        }

        result.SuggestedNextActions = new List<string>
        {
            "Use 'list-tables' to verify record counts",
            "Use 'list-measures' to see updated calculations",
            "Validate Data Model integrity"
        };
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
                result.SuggestedNextActions = new List<string>
                {
                    "Use 'list-measures' to see available measures",
                    "Check measure name for typos",
                    "Verify the Excel file contains a Data Model"
                };
            }
            result.WorkflowHint = $"Failed to delete measure '{measureName}'. Verify measure exists.";
            throw new ModelContextProtocol.McpException($"delete-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Measure '{measureName}' deleted successfully",
                "Use 'list-measures' to verify deletion",
                "Changes saved to workbook"
            };
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
                result.SuggestedNextActions = new List<string>
                {
                    "Use 'list-relationships' to see available relationships",
                    "Check table and column names for typos",
                    "Verify the Excel file contains a Data Model"
                };
            }
            result.WorkflowHint = $"Failed to delete relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Verify relationship exists.";
            throw new ModelContextProtocol.McpException($"delete-relationship failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} deleted successfully",
                "Use 'list-relationships' to verify deletion",
                "Changes saved to workbook"
            };
        }
        result.WorkflowHint = "Relationship deleted. Next, verify remaining relationships or create new ones.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // Phase 2 COM API Action Handlers (Office 2016+)

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
                result.SuggestedNextActions = new List<string>
                {
                    "Use 'list-tables' to see available tables",
                    "Check table name for typos",
                    "Verify the Excel file contains a Data Model"
                };
            }
            result.WorkflowHint = $"Failed to list columns for table '{tableName}'. Verify table exists.";
            throw new ModelContextProtocol.McpException($"list-columns failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'view-table' to see table details",
                "Use 'create-measure' to create calculated measures",
                "Use 'create-relationship' to link tables"
            };
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
                result.SuggestedNextActions = new List<string>
                {
                    "Use 'list-tables' to see available tables",
                    "Check table name for typos",
                    "Verify the Excel file contains a Data Model"
                };
            }
            result.WorkflowHint = $"Failed to view table '{tableName}'. Verify table exists.";
            throw new ModelContextProtocol.McpException($"view-table failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'list-columns' to see all columns in detail",
                "Use 'create-measure' to add calculated measures",
                "Use 'list-relationships' to see table connections"
            };
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
                result.SuggestedNextActions = new List<string>
                {
                    "Verify the Excel file contains a Data Model",
                    "Try 'list-tables' to check if tables exist",
                    "Check if file is corrupted"
                };
            }
            result.WorkflowHint = "Failed to get Data Model info. Verify workbook has a Data Model.";
            throw new ModelContextProtocol.McpException($"get-model-info failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'list-tables' to explore tables",
                "Use 'list-measures' to see calculated measures",
                "Use 'list-relationships' to understand table connections"
            };
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
                result.SuggestedNextActions = new List<string>
                {
                    "Verify DAX formula syntax is correct",
                    "Use 'list-tables' to check available tables",
                    "Use 'list-columns' to see available columns",
                    "Check if measure name already exists"
                };
            }
            result.WorkflowHint = $"Failed to create measure '{measureName}'. Check DAX syntax and table existence.";
            throw new ModelContextProtocol.McpException($"create-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Measure '{measureName}' created successfully in table '{tableName}'",
                "Use 'view-measure' to verify the formula",
                "Use 'list-measures' to see all measures",
                "Changes saved to workbook"
            };
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
                result.SuggestedNextActions = new List<string>
                {
                    "Verify measure name exists using 'list-measures'",
                    "Check DAX formula syntax if updating formula",
                    "Use 'view-measure' to see current formula",
                    "Ensure at least one property is provided for update"
                };
            }
            result.WorkflowHint = $"Failed to update measure '{measureName}'. Verify measure exists and DAX syntax.";
            throw new ModelContextProtocol.McpException($"update-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Measure '{measureName}' updated successfully",
                "Use 'view-measure' to verify the changes",
                "Use 'list-measures' to see all measures",
                "Changes saved to workbook"
            };
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
                result.SuggestedNextActions = new List<string>
                {
                    "Use 'list-tables' to verify both tables exist",
                    "Use 'list-columns' to verify columns exist in both tables",
                    "Check if relationship already exists",
                    "Verify column data types are compatible"
                };
            }
            result.WorkflowHint = $"Failed to create relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Verify tables and columns.";
            throw new ModelContextProtocol.McpException($"create-relationship failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Relationship created from {fromTable}.{fromColumn} to {toTable}.{toColumn}",
                "Use 'list-relationships' to verify the relationship",
                "Create measures that use this relationship",
                "Changes saved to workbook"
            };
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
                result.SuggestedNextActions = new List<string>
                {
                    "Use 'list-relationships' to verify relationship exists",
                    "Check table and column names for typos",
                    "Verify the Excel file contains a Data Model"
                };
            }
            result.WorkflowHint = $"Failed to update relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Verify relationship exists.";
            throw new ModelContextProtocol.McpException($"update-relationship failed for '{filePath}': {result.ErrorMessage}");
        }

        // Success - add workflow guidance
        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} updated successfully",
                "Use 'list-relationships' to verify the changes",
                "Changes saved to workbook"
            };
        }
        result.WorkflowHint = "Relationship updated. Changes will affect measures using this relationship.";

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    // TOM API Action Handlers (Phase 4)

    private static async Task<string> CreateMeasureAsync(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? measureName,
        string? daxFormula,
        string? description,
        string? formatString,
        string? batchId)
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

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.CreateMeasure(filePath, tableName, measureName, daxFormula, description, formatString));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    "Verify table name is correct",
                    "Check DAX formula syntax",
                    "Ensure TOM API connection is available"
                };
            }
            throw new ModelContextProtocol.McpException($"create-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateMeasureAsync(
        DataModelTomCommands commands,
        string filePath,
        string? measureName,
        string? daxFormula,
        string? description,
        string? formatString,
        string? batchId)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'measureName' is required for update-measure action");
        }

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.UpdateMeasure(filePath, measureName, daxFormula, description, formatString));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    "Use 'list-measures' to see available measures",
                    "Verify measure name is correct",
                    "Check DAX formula syntax if updating formula"
                };
            }
            throw new ModelContextProtocol.McpException($"update-measure failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateRelationshipAsync(
        DataModelTomCommands commands,
        string filePath,
        string? fromTable,
        string? fromColumn,
        string? toTable,
        string? toColumn,
        bool? isActive,
        string? crossFilterDirection,
        string? batchId)
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

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.CreateRelationship(
            filePath,
            fromTable,
            fromColumn,
            toTable,
            toColumn,
            isActive ?? true,
            crossFilterDirection ?? "Single"
        ));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    "Verify table and column names",
                    "Check that columns have compatible data types",
                    "Use 'list-tables' to see available tables"
                };
            }
            throw new ModelContextProtocol.McpException($"create-relationship failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateRelationshipAsync(
        DataModelTomCommands commands,
        string filePath,
        string? fromTable,
        string? fromColumn,
        string? toTable,
        string? toColumn,
        bool? isActive,
        string? crossFilterDirection,
        string? batchId)
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

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.UpdateRelationship(
            filePath,
            fromTable,
            fromColumn,
            toTable,
            toColumn,
            isActive,
            crossFilterDirection
        ));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    "Use 'list-relationships' to see available relationships",
                    "Verify relationship exists",
                    "Check table and column names"
                };
            }
            throw new ModelContextProtocol.McpException($"update-relationship failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> CreateCalculatedColumnAsync(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? columnName,
        string? daxFormula,
        string? description,
        string? dataType,
        string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for create-column action");
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'columnName' is required for create-column action");
        }

        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            throw new ModelContextProtocol.McpException("Parameter 'daxFormula' is required for create-column action");
        }

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.CreateCalculatedColumn(
            filePath,
            tableName,
            columnName,
            daxFormula,
            description,
            dataType ?? "String"
        ));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    "Verify table name is correct",
                    "Check DAX formula syntax",
                    "Ensure data type is valid (String, Integer, Double, Boolean, DateTime)"
                };
            }
            throw new ModelContextProtocol.McpException($"create-column failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ValidateDaxAsync(
        DataModelTomCommands commands,
        string filePath,
        string? daxFormula,
        string? batchId)
    {
        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            throw new ModelContextProtocol.McpException("Parameter 'daxFormula' is required for validate-dax action");
        }

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.ValidateDax(filePath, daxFormula));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"validate-dax failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListCalculatedColumnsAsync(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? batchId)
    {
        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.ListCalculatedColumns(filePath, tableName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    "Verify file has Data Model enabled",
                    "Use 'list-tables' to see available tables",
                    "Ensure TOM API connection is available"
                };
            }
            throw new ModelContextProtocol.McpException($"list-columns failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ViewCalculatedColumnAsync(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? columnName,
        string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for view-column action");
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'columnName' is required for view-column action");
        }

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.ViewCalculatedColumn(filePath, tableName, columnName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    $"Use 'list-columns' to see columns in table '{tableName}'",
                    "Check column name spelling",
                    "Verify table exists in Data Model"
                };
            }
            throw new ModelContextProtocol.McpException($"view-column failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> UpdateCalculatedColumnAsync(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? columnName,
        string? daxFormula,
        string? description,
        string? dataType,
        string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for update-column action");
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'columnName' is required for update-column action");
        }

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.UpdateCalculatedColumn(
            filePath,
            tableName,
            columnName,
            daxFormula,
            description,
            dataType
        ));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    $"Use 'list-columns' to see columns in table '{tableName}'",
                    "Verify column exists",
                    "Check DAX formula syntax if updating formula"
                };
            }
            throw new ModelContextProtocol.McpException($"update-column failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> DeleteCalculatedColumnAsync(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? columnName,
        string? batchId)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for delete-column action");
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'columnName' is required for delete-column action");
        }

        // TOM API doesn't use Excel COM batching - it uses its own Analysis Services connection
        var result = await Task.Run(() => commands.DeleteCalculatedColumn(filePath, tableName, columnName));

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
            {
                result.SuggestedNextActions = new List<string>
                {
                    $"Use 'list-columns' to see columns in table '{tableName}'",
                    "Verify column exists",
                    "Check that column is not referenced by other objects"
                };
            }
            throw new ModelContextProtocol.McpException($"delete-column failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
