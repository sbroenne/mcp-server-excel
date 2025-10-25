using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json;
using ModelContextProtocol.Server;
using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel Table (ListObject) management tool for MCP server.
/// Handles creating, listing, renaming, and deleting Excel Tables for Power Query integration.
///
/// LLM Usage Patterns:
/// - Use "list" to see all Excel Tables in a workbook
/// - Use "create" to convert ranges to Excel Tables (enables Power Query references)
/// - Use "info" to get detailed information about a table
/// - Use "rename" to change table names (update Power Query references accordingly)
/// - Use "delete" to remove tables (converts back to range, data preserved)
///
/// IMPORTANT:
/// - Excel Tables are the recommended way to reference data in Power Query
/// - Power Query syntax: Excel.CurrentWorkbook(){[Name="TableName"]}[Content]
/// - Table names must start with a letter/underscore, contain only alphanumeric and underscore characters
/// - Deleting a table converts it back to a range but preserves data
/// </summary>
[McpServerToolType]
public static class TableTool
{
    /// <summary>
    /// Manage Excel Tables (ListObjects) - comprehensive table management including Power Pivot integration
    /// </summary>
    [McpServerTool(Name = "table")]
    [Description("Manage Excel Tables (ListObjects) for Power Query integration. Supports: list, create, info, rename, delete, resize, toggle-totals, set-column-total, read, append, set-style, add-to-datamodel.")]
    public static string Table(
        [Required]
        [RegularExpression("^(list|create|info|rename|delete|resize|toggle-totals|set-column-total|read|append|set-style|add-to-datamodel)$")]
        [Description("Action: list, create, info, rename, delete, resize, toggle-totals, set-column-total, read, append, set-style, add-to-datamodel")]
        string action,

        [Required]
        [FileExtensions(Extensions = "xlsx,xlsm")]
        [Description("Excel file path (.xlsx or .xlsm)")]
        string excelPath,

        [StringLength(255, MinimumLength = 1)]
        [RegularExpression(@"^[a-zA-Z_][a-zA-Z0-9_]*$")]
        [Description("Table name (required for most actions). Must start with letter/underscore, alphanumeric + underscore only")]
        string? tableName = null,

        [StringLength(31, MinimumLength = 1)]
        [RegularExpression(@"^[^[\]/*?\\:]+$")]
        [Description("Sheet name (required for create)")]
        string? sheetName = null,

        [Description("Excel range (e.g., 'A1:D10') - required for create/resize")]
        string? range = null,

        [StringLength(255, MinimumLength = 1)]
        [RegularExpression(@"^[a-zA-Z_][a-zA-Z0-9_]*$")]
        [Description("New table name (required for rename) or column name (required for set-column-total)")]
        string? newName = null,

        [Description("Whether the range has headers (default: true for create) or show totals (for toggle-totals)")]
        bool hasHeaders = true,

        [Description("Table style name (e.g., 'TableStyleMedium2') for create/set-style, or total function (sum/avg/count) for set-column-total, or CSV data for append")]
        string? tableStyle = null)
    {
        try
        {
            var tableCommands = new TableCommands();

            return action.ToLowerInvariant() switch
            {
                "list" => ListTables(tableCommands, excelPath),
                "create" => CreateTable(tableCommands, excelPath, sheetName, tableName, range, hasHeaders, tableStyle),
                "info" => GetTableInfo(tableCommands, excelPath, tableName),
                "rename" => RenameTable(tableCommands, excelPath, tableName, newName),
                "delete" => DeleteTable(tableCommands, excelPath, tableName),
                "resize" => ResizeTable(tableCommands, excelPath, tableName, range),
                "toggle-totals" => ToggleTotals(tableCommands, excelPath, tableName, hasHeaders),
                "set-column-total" => SetColumnTotal(tableCommands, excelPath, tableName, newName, tableStyle),
                "read" => ReadTableData(tableCommands, excelPath, tableName),
                "append" => AppendRows(tableCommands, excelPath, tableName, tableStyle),
                "set-style" => SetTableStyle(tableCommands, excelPath, tableName, tableStyle),
                "add-to-datamodel" => AddToDataModel(tableCommands, excelPath, tableName),
                _ => throw new ModelContextProtocol.McpException(
                    $"Unknown action '{action}'. Supported: list, create, info, rename, delete, resize, toggle-totals, set-column-total, read, append, set-style, add-to-datamodel")
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

    private static string ListTables(TableCommands commands, string filePath)
    {
        var result = commands.List(filePath);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the Excel file exists and is accessible",
                "Verify the file path is correct"
            };
            result.WorkflowHint = "List failed. Ensure the file exists and retry.";
            throw new ModelContextProtocol.McpException($"list failed for '{filePath}': {result.ErrorMessage}");
        }

        if (result.Tables == null || !result.Tables.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'table create' to create an Excel Table from a range",
                "Excel Tables enable Power Query references: Excel.CurrentWorkbook(){[Name=\"TableName\"]}[Content]",
                "Tables provide auto-filtering, structured references, and dynamic expansion"
            };
            result.WorkflowHint = "No tables found. Create tables to enable Power Query integration.";
        }
        else
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'table info <tableName>' to view detailed table information",
                "Reference tables in Power Query: Excel.CurrentWorkbook(){[Name=\"TableName\"]}[Content]",
                "Use 'table rename <oldName> <newName>' to rename a table"
            };
            result.WorkflowHint = $"Found {result.Tables.Count} table(s). Use 'table info' for details.";
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string CreateTable(TableCommands commands, string filePath, string? sheetName, string? tableName, string? range, bool hasHeaders, string? tableStyle)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) ExcelToolsBase.ThrowMissingParameter(nameof(sheetName), "create");
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "create");
        if (string.IsNullOrWhiteSpace(range)) ExcelToolsBase.ThrowMissingParameter(nameof(range), "create");

        var result = commands.Create(filePath, sheetName!, tableName!, range!, hasHeaders, tableStyle);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Check that the sheet name exists in the workbook",
                "Verify the range is valid (e.g., 'A1:D10')",
                "Ensure the table name is unique and follows naming rules (starts with letter/underscore, alphanumeric + underscore only)",
                "Check that the range contains data"
            };
            result.WorkflowHint = "Table creation failed. Verify sheet name, range, and table name.";
            throw new ModelContextProtocol.McpException($"create failed for table '{tableName}': {result.ErrorMessage}");
        }

        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Use 'table info {tableName}' to view table details",
                $"Reference in Power Query: Excel.CurrentWorkbook(){{[Name=\"{tableName}\"]}}[Content]",
                $"Use 'table rename {tableName} NewName' to rename the table"
            };
        }

        if (string.IsNullOrEmpty(result.WorkflowHint))
        {
            result.WorkflowHint = $"Table '{tableName}' created successfully. Ready for Power Query integration.";
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string GetTableInfo(TableCommands commands, string filePath, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "info");

        var result = commands.GetInfo(filePath, tableName!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'table list' to see all available tables",
                "Check that the table name is correct (names are case-sensitive)",
                "Verify the Excel file exists and is accessible"
            };
            result.WorkflowHint = "Table not found. Use 'table list' to see available tables.";
            throw new ModelContextProtocol.McpException($"info failed for table '{tableName}': {result.ErrorMessage}");
        }

        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Use 'table rename {tableName} NewName' to rename the table",
                $"Use 'table delete {tableName}' to remove the table (data preserved as range)",
                $"Reference in Power Query: Excel.CurrentWorkbook(){{[Name=\"{tableName}\"]}}[Content]"
            };
        }

        if (string.IsNullOrEmpty(result.WorkflowHint))
        {
            result.WorkflowHint = $"Table '{tableName}' details retrieved. {result.Table?.RowCount ?? 0} rows, {result.Table?.ColumnCount ?? 0} columns.";
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string RenameTable(TableCommands commands, string filePath, string? tableName, string? newName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "rename");
        if (string.IsNullOrWhiteSpace(newName)) ExcelToolsBase.ThrowMissingParameter(nameof(newName), "rename");

        var result = commands.Rename(filePath, tableName!, newName!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'table list' to see all available tables",
                "Check that the table name is correct",
                "Ensure the new name is unique and follows naming rules (starts with letter/underscore, alphanumeric + underscore only)",
                "Verify the Excel file is not open in Excel Desktop"
            };
            result.WorkflowHint = "Rename failed. Verify table name and new name are valid.";
            throw new ModelContextProtocol.McpException($"rename failed for table '{tableName}': {result.ErrorMessage}");
        }

        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                $"Update Power Query references to use new name: Excel.CurrentWorkbook(){{[Name=\"{newName}\"]}}[Content]",
                "Update any formulas or scripts that reference the old table name",
                $"Use 'table info {newName}' to verify the rename"
            };
        }

        if (string.IsNullOrEmpty(result.WorkflowHint))
        {
            result.WorkflowHint = $"Table renamed from '{tableName}' to '{newName}'. Update Power Query references.";
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string DeleteTable(TableCommands commands, string filePath, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "delete");

        var result = commands.Delete(filePath, tableName!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            result.SuggestedNextActions = new List<string>
            {
                "Use 'table list' to see all available tables",
                "Check that the table name is correct",
                "Verify the Excel file is not open in Excel Desktop"
            };
            result.WorkflowHint = "Delete failed. Verify table name is correct.";
            throw new ModelContextProtocol.McpException($"delete failed for table '{tableName}': {result.ErrorMessage}");
        }

        if (result.SuggestedNextActions == null || !result.SuggestedNextActions.Any())
        {
            result.SuggestedNextActions = new List<string>
            {
                "Data has been preserved as a regular range",
                "Update or remove Power Query expressions that referenced this table",
                "Use 'worksheet read' to access the data as a range"
            };
        }

        if (string.IsNullOrEmpty(result.WorkflowHint))
        {
            result.WorkflowHint = $"Table '{tableName}' deleted. Data converted back to regular range.";
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ResizeTable(TableCommands commands, string filePath, string? tableName, string? newRange)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "resize");
        if (string.IsNullOrWhiteSpace(newRange)) ExcelToolsBase.ThrowMissingParameter(nameof(newRange), "resize");

        var result = commands.Resize(filePath, tableName!, newRange!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"resize failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ToggleTotals(TableCommands commands, string filePath, string? tableName, bool showTotals)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "toggle-totals");

        var result = commands.ToggleTotals(filePath, tableName!, showTotals);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"toggle-totals failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetColumnTotal(TableCommands commands, string filePath, string? tableName, string? columnName, string? totalFunction)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-column-total");
        if (string.IsNullOrWhiteSpace(columnName)) ExcelToolsBase.ThrowMissingParameter(nameof(columnName), "set-column-total");
        if (string.IsNullOrWhiteSpace(totalFunction)) ExcelToolsBase.ThrowMissingParameter(nameof(totalFunction), "set-column-total");

        var result = commands.SetColumnTotal(filePath, tableName!, columnName!, totalFunction!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-column-total failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string ReadTableData(TableCommands commands, string filePath, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "read");

        var result = commands.ReadData(filePath, tableName!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"read failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string AppendRows(TableCommands commands, string filePath, string? tableName, string? csvData)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "append");
        if (string.IsNullOrWhiteSpace(csvData)) ExcelToolsBase.ThrowMissingParameter(nameof(csvData), "append");

        var result = commands.AppendRows(filePath, tableName!, csvData!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"append failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string SetTableStyle(TableCommands commands, string filePath, string? tableName, string? tableStyle)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "set-style");
        if (string.IsNullOrWhiteSpace(tableStyle)) ExcelToolsBase.ThrowMissingParameter(nameof(tableStyle), "set-style");

        var result = commands.SetStyle(filePath, tableName!, tableStyle!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"set-style failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }

    private static string AddToDataModel(TableCommands commands, string filePath, string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) ExcelToolsBase.ThrowMissingParameter(nameof(tableName), "add-to-datamodel");

        var result = commands.AddToDataModel(filePath, tableName!);

        if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
        {
            throw new ModelContextProtocol.McpException($"add-to-datamodel failed for table '{tableName}': {result.ErrorMessage}");
        }

        return JsonSerializer.Serialize(result, ExcelToolsBase.JsonOptions);
    }
}
