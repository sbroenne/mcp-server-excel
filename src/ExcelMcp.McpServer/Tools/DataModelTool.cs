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
public static class DataModelTool
{
    /// <summary>
    /// Manage Excel Data Model (Power Pivot) - tables, measures, relationships
    /// </summary>
    [McpServerTool(Name = "datamodel")]
    [Description("Manage Excel Data Model operations. Supports: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship, create-measure, update-measure, create-relationship, update-relationship, create-column, list-columns, view-column, update-column, delete-column, validate-dax.")]
    public static async Task<string> DataModel(
        [Required]
        [RegularExpression("^(list-tables|list-measures|view-measure|export-measure|list-relationships|refresh|delete-measure|delete-relationship|create-measure|update-measure|create-relationship|update-relationship|create-column|list-columns|view-column|update-column|delete-column|validate-dax)$")]
        [Description("Action: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship, create-measure, update-measure, create-relationship, update-relationship, create-column, list-columns, view-column, update-column, delete-column, validate-dax")]
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
        string? dataType = null)
    {
        try
        {
            var dataModelCommands = new DataModelCommands();
            var tomCommands = new DataModelTomCommands();

            return action.ToLowerInvariant() switch
            {
                // COM API operations (Phase 1)
                "list-tables" => ListTables(dataModelCommands, excelPath),
                "list-measures" => ListMeasures(dataModelCommands, excelPath),
                "view-measure" => ViewMeasure(dataModelCommands, excelPath, measureName),
                "export-measure" => await ExportMeasure(dataModelCommands, excelPath, measureName, outputPath),
                "list-relationships" => ListRelationships(dataModelCommands, excelPath),
                "refresh" => Refresh(dataModelCommands, excelPath),
                "delete-measure" => DeleteMeasure(dataModelCommands, excelPath, measureName),
                "delete-relationship" => DeleteRelationship(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn),

                // TOM API operations (Phase 4)
                "create-measure" => CreateMeasure(tomCommands, excelPath, tableName, measureName, daxFormula, description, formatString),
                "update-measure" => UpdateMeasure(tomCommands, excelPath, measureName, daxFormula, description, formatString),
                "create-relationship" => CreateRelationship(tomCommands, excelPath, fromTable, fromColumn, toTable, toColumn, isActive, crossFilterDirection),
                "update-relationship" => UpdateRelationship(tomCommands, excelPath, fromTable, fromColumn, toTable, toColumn, isActive, crossFilterDirection),
                "create-column" => CreateCalculatedColumn(tomCommands, excelPath, tableName, columnName, daxFormula, description, dataType),
                "list-columns" => ListCalculatedColumns(tomCommands, excelPath, tableName),
                "view-column" => ViewCalculatedColumn(tomCommands, excelPath, tableName, columnName),
                "update-column" => UpdateCalculatedColumn(tomCommands, excelPath, tableName, columnName, daxFormula, description, dataType),
                "delete-column" => DeleteCalculatedColumn(tomCommands, excelPath, tableName, columnName),
                "validate-dax" => ValidateDax(tomCommands, excelPath, daxFormula),

                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship, create-measure, update-measure, create-relationship, update-relationship, create-column, list-columns, view-column, update-column, delete-column, validate-dax")
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

    private static string ListTables(DataModelCommands commands, string filePath)
    {
        var result = commands.ListTables(filePath);

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

    private static string ListMeasures(DataModelCommands commands, string filePath)
    {
        var result = commands.ListMeasures(filePath);

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

    private static string ViewMeasure(DataModelCommands commands, string filePath, string? measureName)
    {
        if (string.IsNullOrEmpty(measureName))
            throw new ModelContextProtocol.McpException("measureName is required for view-measure action");

        var result = commands.ViewMeasure(filePath, measureName);

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

    private static async Task<string> ExportMeasure(DataModelCommands commands, string filePath, string? measureName, string? outputPath)
    {
        if (string.IsNullOrEmpty(measureName))
            throw new ModelContextProtocol.McpException("measureName is required for export-measure action");

        if (string.IsNullOrEmpty(outputPath))
            throw new ModelContextProtocol.McpException("outputPath is required for export-measure action");

        var result = await commands.ExportMeasure(filePath, measureName, outputPath);

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

    private static string ListRelationships(DataModelCommands commands, string filePath)
    {
        var result = commands.ListRelationships(filePath);

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

    private static string Refresh(DataModelCommands commands, string filePath)
    {
        var result = commands.Refresh(filePath);

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

    private static string DeleteMeasure(DataModelCommands commands, string filePath, string? measureName)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'measureName' is required for delete-measure action");
        }

        var result = commands.DeleteMeasure(filePath, measureName);

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

    private static string DeleteRelationship(DataModelCommands commands, string filePath,
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

        var result = commands.DeleteRelationship(filePath, fromTable, fromColumn, toTable, toColumn);

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

    // TOM API Action Handlers (Phase 4)

    private static string CreateMeasure(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? measureName,
        string? daxFormula,
        string? description,
        string? formatString)
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

        var result = commands.CreateMeasure(filePath, tableName, measureName, daxFormula, description, formatString);

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

    private static string UpdateMeasure(
        DataModelTomCommands commands,
        string filePath,
        string? measureName,
        string? daxFormula,
        string? description,
        string? formatString)
    {
        if (string.IsNullOrWhiteSpace(measureName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'measureName' is required for update-measure action");
        }

        var result = commands.UpdateMeasure(filePath, measureName, daxFormula, description, formatString);

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

    private static string CreateRelationship(
        DataModelTomCommands commands,
        string filePath,
        string? fromTable,
        string? fromColumn,
        string? toTable,
        string? toColumn,
        bool? isActive,
        string? crossFilterDirection)
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

        var result = commands.CreateRelationship(
            filePath,
            fromTable,
            fromColumn,
            toTable,
            toColumn,
            isActive ?? true,
            crossFilterDirection ?? "Single"
        );

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

    private static string UpdateRelationship(
        DataModelTomCommands commands,
        string filePath,
        string? fromTable,
        string? fromColumn,
        string? toTable,
        string? toColumn,
        bool? isActive,
        string? crossFilterDirection)
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

        var result = commands.UpdateRelationship(
            filePath,
            fromTable,
            fromColumn,
            toTable,
            toColumn,
            isActive,
            crossFilterDirection
        );

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

    private static string CreateCalculatedColumn(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? columnName,
        string? daxFormula,
        string? description,
        string? dataType)
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

        var result = commands.CreateCalculatedColumn(
            filePath,
            tableName,
            columnName,
            daxFormula,
            description,
            dataType ?? "String"
        );

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

    private static string ValidateDax(
        DataModelTomCommands commands,
        string filePath,
        string? daxFormula)
    {
        if (string.IsNullOrWhiteSpace(daxFormula))
        {
            throw new ModelContextProtocol.McpException("Parameter 'daxFormula' is required for validate-dax action");
        }

        var result = commands.ValidateDax(filePath, daxFormula);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"validate-dax failed for '{filePath}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ListCalculatedColumns(
        DataModelTomCommands commands,
        string filePath,
        string? tableName)
    {
        var result = commands.ListCalculatedColumns(filePath, tableName);

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

    private static string ViewCalculatedColumn(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for view-column action");
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'columnName' is required for view-column action");
        }

        var result = commands.ViewCalculatedColumn(filePath, tableName, columnName);

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

    private static string UpdateCalculatedColumn(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? columnName,
        string? daxFormula,
        string? description,
        string? dataType)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for update-column action");
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'columnName' is required for update-column action");
        }

        var result = commands.UpdateCalculatedColumn(
            filePath,
            tableName,
            columnName,
            daxFormula,
            description,
            dataType
        );

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

    private static string DeleteCalculatedColumn(
        DataModelTomCommands commands,
        string filePath,
        string? tableName,
        string? columnName)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'tableName' is required for delete-column action");
        }

        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ModelContextProtocol.McpException("Parameter 'columnName' is required for delete-column action");
        }

        var result = commands.DeleteCalculatedColumn(filePath, tableName, columnName);

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
