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
///
/// Phase 1 Scope (COM API):
/// - Tables: List and view metadata
/// - Measures: List, view, and export DAX formulas
/// - Relationships: List relationships between tables
/// - Refresh: Refresh Data Model connections
///
/// Out of Scope (requires TOM API - Phase 4):
/// - Calculated columns
/// - Advanced DAX operations
/// </summary>
[McpServerToolType]
public static class ExcelDataModelTool
{
    /// <summary>
    /// Manage Excel Data Model (Power Pivot) - tables, measures, relationships
    /// </summary>
    [McpServerTool(Name = "excel_datamodel")]
    [Description("Manage Excel Data Model operations. Supports: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship.")]
    public static async Task<string> ExcelDataModel(
        [Required]
        [RegularExpression("^(list-tables|list-measures|view-measure|export-measure|list-relationships|refresh|delete-measure|delete-relationship)$")]
        [Description("Action: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship")]
        string action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(255, MinimumLength = 1)]
        [Description("Measure name (for view-measure, export-measure, delete-measure)")]
        string? measureName = null,

        [FileExtensions(Extensions = "dax")]
        [Description("Output file path for DAX export (for export-measure)")]
        string? outputPath = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Source table name (for delete-relationship)")]
        string? fromTable = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Source column name (for delete-relationship)")]
        string? fromColumn = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Target table name (for delete-relationship)")]
        string? toTable = null,

        [StringLength(255, MinimumLength = 1)]
        [Description("Target column name (for delete-relationship)")]
        string? toColumn = null)
    {
        try
        {
            var dataModelCommands = new DataModelCommands();

            return action.ToLowerInvariant() switch
            {
                "list-tables" => ListTables(dataModelCommands, excelPath),
                "list-measures" => ListMeasures(dataModelCommands, excelPath),
                "view-measure" => ViewMeasure(dataModelCommands, excelPath, measureName),
                "export-measure" => await ExportMeasure(dataModelCommands, excelPath, measureName, outputPath),
                "list-relationships" => ListRelationships(dataModelCommands, excelPath),
                "refresh" => Refresh(dataModelCommands, excelPath),
                "delete-measure" => DeleteMeasure(dataModelCommands, excelPath, measureName),
                "delete-relationship" => DeleteRelationship(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list-tables, list-measures, view-measure, export-measure, list-relationships, refresh, delete-measure, delete-relationship")
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
}
