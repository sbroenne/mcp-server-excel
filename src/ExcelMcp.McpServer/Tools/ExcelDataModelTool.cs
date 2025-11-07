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

        [Description("Timeout in minutes for data model operations. Default: 2 minutes")]
        double? timeout = null,

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
                DataModelAction.Refresh => await RefreshAsync(dataModelCommands, excelPath, timeout, batchId),
                DataModelAction.DeleteMeasure => await DeleteMeasureAsync(dataModelCommands, excelPath, measureName, batchId),
                DataModelAction.DeleteRelationship => await DeleteRelationshipAsync(dataModelCommands, excelPath, fromTable, fromColumn, toTable, toColumn, batchId),
                DataModelAction.GetTable => await ViewTableAsync(dataModelCommands, excelPath, tableName, batchId),
                DataModelAction.ListColumns => await ListColumnsAsync(dataModelCommands, excelPath, tableName, batchId),
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Tables,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {result.Tables.Count} Data Model tables. Review structure and relationships for analytics."
                : "Failed to list Data Model tables. Verify workbook contains Power Pivot data.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'view-table' to see detailed table information", "Use 'list-measures' to view DAX calculations", "Use 'list-relationships' to see table connections" }
                : ["Verify workbook has Data Model enabled", "Check if tables loaded via Power Query or manual import", "Use excel_powerquery list to see available queries"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListMeasuresAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListMeasuresAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Measures,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {result.Measures.Count} DAX measures in Data Model. Review formulas and table assignments."
                : "Failed to list measures. Verify Data Model contains DAX measures.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'view-measure' to inspect specific DAX formulas", "Use 'export-measure' to save DAX for version control", "Use 'create-measure' to add new calculations" }
                : ["Verify Data Model is properly configured", "Use 'list-tables' to see available tables", "Create measures via 'create-measure' action"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.MeasureName,
            result.DaxFormula,
            result.TableName,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Retrieved DAX formula for '{measureName}'. Review calculation logic and dependencies."
                : $"Failed to view measure '{measureName}'. Verify measure exists in Data Model.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'export-measure' to save DAX for documentation", "Use 'update-measure' to modify formula or format", "Use 'list-measures' to see all available measures" }
                : ["Use 'list-measures' to find correct measure name", "Check for typos in measure name", "Verify Data Model is loaded"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.FilePath,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Exported measure '{measureName}' to '{outputPath}'. Store in version control for DAX management."
                : $"Failed to export measure '{measureName}'. Verify measure exists and output path is writable.",
            suggestedNextActions = result.Success
                ? new[] { "Commit .dax file to version control", "Share DAX formulas with team", "Use as template for similar measures" }
                : ["Use 'view-measure' to verify measure exists", "Check directory permissions for output path", "Ensure parent directory exists"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListRelationshipsAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.ListRelationshipsAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.Relationships,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Found {result.Relationships.Count} relationships in Data Model. Review table connections and cardinality."
                : "Failed to retrieve relationships. Verify Data Model is loaded.",
            suggestedNextActions = result.Success
                ? new[] { "Review relationship directions (one-to-many, many-to-one)", "Verify active relationships for DAX calculations", "Use 'create-relationship' to add connections" }
                : ["Use 'list-tables' to verify tables exist", "Check if Data Model is properly loaded", "Ensure workbook has Power Pivot enabled"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> RefreshAsync(DataModelCommands commands, string filePath, double? timeoutMinutes, string? batchId)
    {
        try
        {
            var timeoutSpan = timeoutMinutes.HasValue ? (TimeSpan?)TimeSpan.FromMinutes(timeoutMinutes.Value) : null;
            var result = await ExcelToolsBase.WithBatchAsync(
                batchId,
                filePath,
                save: true,
                async (batch) => await commands.RefreshAsync(batch, null, timeoutSpan));

            return JsonSerializer.Serialize(new
            {
                result.Success,
                result.ErrorMessage,
                workflowHint = result.Success
                    ? "Data Model refreshed successfully. All tables reloaded from source connections."
                    : $"Data Model refresh failed. {result.ErrorMessage}",
                suggestedNextActions = result.Success
                    ? new[] { "Verify data with 'list-tables' (check record counts)", "Test DAX measures with updated data", "Use 'get-model-info' for refresh summary" }
                    : ["Check connection credentials and connectivity", "Use 'list-tables' to identify failing tables", "Review error message for specific data source issues"]
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
                result.RetryGuidance,

                // Workflow hints - MCP Server layer responsibility
                WorkflowHint = "Data Model refresh timeout - check for blocking dialogs or data source issues",
                SuggestedNextActions = new[]
                {
                    "Check if Excel is showing a dialog or is unresponsive",
                    "Verify all data source connections in the Data Model are accessible",
                    "For large Data Models (millions of rows), refresh may genuinely require 5+ minutes",
                    "Consider refreshing individual tables instead of entire model (use tableName parameter)"
                }
            };

            return JsonSerializer.Serialize(response, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Deleted DAX measure '{measureName}' from Data Model. Changes saved to workbook."
                : $"Failed to delete measure '{measureName}'. Verify measure exists.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list-measures' to verify deletion", "Update dependent DAX calculations if needed", "Remove measure references from PivotTables" }
                : ["Use 'list-measures' to find correct measure name", "Check for typos in measure name", "Verify Data Model is loaded"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Deleted relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Changes saved to workbook."
                : $"Failed to delete relationship. Verify relationship exists.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list-relationships' to verify deletion", "Update DAX formulas that relied on this relationship", "Consider creating alternative relationship paths" }
                : ["Use 'list-relationships' to find correct relationship", "Check table and column names for typos", "Verify Data Model is loaded"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableName,
            result.SourceName,
            result.RecordCount,
            result.Columns,
            result.MeasureCount,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Table '{tableName}' has {result.RecordCount:N0} rows, {result.Columns.Count} columns, {result.MeasureCount} measures. Source: {result.SourceName}"
                : $"Failed to view table '{tableName}'. Verify table exists in Data Model.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list-columns' for detailed column information", "Review source query or connection", "Check measure definitions if count > 0" }
                : ["Use 'list-tables' to find correct table name", "Verify Data Model is loaded", "Check for typos in table name"]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> ListColumnsAsync(DataModelCommands commands, string filePath,
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

        // Add workflow hints
        var inBatch = !string.IsNullOrEmpty(batchId);
        var columnCount = result.Columns?.Count ?? 0;
        var calculatedCount = result.Columns?.Count(c => c.IsCalculated) ?? 0;

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            result.TableName,
            result.Columns,
            workflowHint = result.Success
                ? $"Table '{tableName}' has {columnCount} columns ({calculatedCount} calculated, {columnCount - calculatedCount} regular)."
                : $"Failed to list columns for table '{tableName}': {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[]
                {
                    "Use excel_datamodel 'view-table' to see full table details including measures",
                    "Use excel_datamodel 'create-relationship' to link tables by columns",
                    "Use excel_datamodel 'list-measures' to see DAX calculations on this table",
                    inBatch ? "Query more tables in this batch" : "Exploring multiple tables? Use excel_batch for efficiency"
                }
                :
                [
                    "Check table name is correct with excel_datamodel 'list-tables'",
                    "Verify Data Model exists with excel_datamodel 'get-model-info'",
                    "Review error message for specific issue"
                ]
        }, ExcelToolsBase.JsonOptions);
    }

    private static async Task<string> GetModelInfoAsync(DataModelCommands commands, string filePath, string? batchId)
    {
        var result = await ExcelToolsBase.WithBatchAsync(
            batchId,
            filePath,
            save: false,
            async (batch) => await commands.GetInfoAsync(batch));

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.TableCount,
            result.MeasureCount,
            result.RelationshipCount,
            result.TotalRows,
            result.TableNames,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Data Model contains {result.TableCount} tables, {result.MeasureCount} measures, {result.RelationshipCount} relationships. Total: {result.TotalRows:N0} rows."
                : "Failed to retrieve Data Model information. Verify Data Model exists.",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list-tables' to see individual table details", "Use 'list-relationships' to review table connections", "Use 'list-measures' to see all DAX calculations" }
                : ["Verify workbook has Power Pivot enabled", "Check if Data Model is loaded", "Try opening workbook in Excel first"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Created DAX measure '{measureName}' in table '{tableName}'. Formula: {daxFormula}"
                : $"Failed to create measure '{measureName}'. {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[] { "Use 'view-measure' to verify formula and format", "Test measure in PivotTable or DAX query", "Use 'export-measure' for version control" }
                : ["Verify table name with 'list-tables'", "Check DAX syntax in formula", "Ensure measure name doesn't already exist"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Updated DAX measure '{measureName}'. Changes saved to workbook."
                : $"Failed to update measure '{measureName}'. {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[] { "Use 'view-measure' to verify changes", "Refresh PivotTables using this measure", "Test formula with sample data" }
                : ["Use 'list-measures' to verify measure exists", "Check DAX syntax if formula was changed", "Verify measure name is correct"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Created relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Active: {isActive ?? true}"
                : $"Failed to create relationship. {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list-relationships' to verify creation", "Test DAX formulas using this relationship", "Verify relationship direction (one-to-many)" }
                : ["Check table and column names with 'list-tables' and 'list-columns'", "Verify columns have compatible data types", "Ensure no duplicate relationships exist"]
        }, ExcelToolsBase.JsonOptions);
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

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.ErrorMessage,
            workflowHint = result.Success
                ? $"Updated relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn}. Active: {isActive.Value}"
                : $"Failed to update relationship. {result.ErrorMessage}",
            suggestedNextActions = result.Success
                ? new[] { "Use 'list-relationships' to verify status change", "Test DAX formulas if relationship activation changed", "Verify only one relationship is active between tables" }
                : ["Use 'list-relationships' to find correct relationship", "Check table and column names for typos", "Verify relationship exists before updating"]
        }, ExcelToolsBase.JsonOptions);
    }
}
