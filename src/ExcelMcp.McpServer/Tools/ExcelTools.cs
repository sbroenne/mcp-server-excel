using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;

#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' requirements

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Main Excel tools registry for Model Context Protocol (MCP) server.
/// 
/// This class consolidates all Excel automation tools into a single entry point
/// optimized for LLM usage patterns. Each tool is focused on a specific Excel domain:
/// 
/// 🔧 Tool Architecture:
/// - ExcelFileTool: File operations (create, validate, check existence)
/// - ExcelPowerQueryTool: M code and data loading management  
/// - ExcelWorksheetTool: Sheet operations and bulk data handling
/// - ExcelParameterTool: Named ranges as configuration parameters
/// - ExcelCellTool: Precise individual cell operations
/// - ExcelVbaTool: VBA macro management and execution
/// 
/// 🤖 LLM Usage Guidelines:
/// 1. Start with ExcelFileTool to create or validate files
/// 2. Use ExcelWorksheetTool for data operations and sheet management
/// 3. Use ExcelPowerQueryTool for advanced data transformation
/// 4. Use ExcelParameterTool for configuration and reusable values
/// 5. Use ExcelCellTool for precision operations on individual cells
/// 6. Use ExcelVbaTool for complex automation (requires .xlsm files)
/// 
/// 📝 Parameter Patterns:
/// - action: Always the first parameter, defines what operation to perform
/// - filePath: Excel file path (.xlsx or .xlsm based on requirements)
/// - Context-specific parameters: Each tool has domain-appropriate parameters
/// 
/// 🎯 Design Philosophy:
/// - Resource-based: Tools represent Excel domains, not individual operations
/// - Action-oriented: Each tool supports multiple related actions
/// - LLM-friendly: Clear naming, comprehensive documentation, predictable patterns
/// - Error-consistent: Standardized error handling across all tools
/// </summary>
[McpServerToolType]
[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicMethods)]
public static class ExcelTools
{
    // File Operations
    /// <summary>
    /// Manage Excel files - create, validate, and check file operations
    /// Delegates to ExcelFileTool for implementation.
    /// </summary>
    [McpServerTool(Name = "excel_file")]
    [Description("Create, validate, and manage Excel files (.xlsx, .xlsm). Supports actions: create-empty, validate, check-exists.")]
    public static string ExcelFile(
        [Description("Action to perform: create-empty, validate, check-exists")] string action,
        [Description("Excel file path (.xlsx or .xlsm extension)")] string filePath,
        [Description("Optional: macro-enabled flag for create-empty (default: false)")] bool macroEnabled = false)
        => ExcelFileTool.ExcelFile(action, filePath, macroEnabled);

    // Power Query Operations  
    /// <summary>
    /// Manage Power Query operations - M code, data loading, and query lifecycle
    /// Delegates to ExcelPowerQueryTool for implementation.
    /// </summary>
    [McpServerTool(Name = "excel_powerquery")]
    [Description("Manage Power Query M code and data loading. Supports: list, view, import, export, update, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config.")]
    public static string ExcelPowerQuery(
        [Description("Action: list, view, import, export, update, delete, set-load-to-table, set-load-to-data-model, set-load-to-both, set-connection-only, get-load-config")] string action,
        [Description("Excel file path (.xlsx or .xlsm)")] string filePath,
        [Description("Power Query name (required for most actions)")] string? queryName = null,
        [Description("Source .pq file path (for import/update) or target file path (for export)")] string? sourceOrTargetPath = null,
        [Description("Target worksheet name (for set-load-to-table action)")] string? targetSheet = null)
        => ExcelPowerQueryTool.ExcelPowerQuery(action, filePath, queryName, sourceOrTargetPath, targetSheet);

    // Worksheet Operations
    /// <summary>
    /// Manage Excel worksheets - data operations, sheet management, and content manipulation
    /// Delegates to ExcelWorksheetTool for implementation.
    /// </summary>
    [McpServerTool(Name = "excel_worksheet")]
    [Description("Manage Excel worksheets and data. Supports: list, read, write, create, rename, copy, delete, clear, append.")]
    public static string ExcelWorksheet(
        [Description("Action: list, read, write, create, rename, copy, delete, clear, append")] string action,
        [Description("Excel file path (.xlsx or .xlsm)")] string filePath,
        [Description("Worksheet name (required for most actions)")] string? sheetName = null,
        [Description("Excel range (e.g., 'A1:D10' for read/clear) or CSV file path (for write/append)")] string? range = null,
        [Description("New sheet name (for rename) or source sheet name (for copy)")] string? targetName = null)
        => ExcelWorksheetTool.ExcelWorksheet(action, filePath, sheetName, range, targetName);

    // Parameter Operations
    /// <summary>
    /// Manage Excel parameters (named ranges) - configuration values and reusable references
    /// Delegates to ExcelParameterTool for implementation.
    /// </summary>
    [McpServerTool(Name = "excel_parameter")]
    [Description("Manage Excel named ranges as parameters. Supports: list, get, set, create, delete.")]
    public static string ExcelParameter(
        [Description("Action: list, get, set, create, delete")] string action,
        [Description("Excel file path (.xlsx or .xlsm)")] string filePath,
        [Description("Parameter (named range) name")] string? parameterName = null,
        [Description("Parameter value (for set) or cell reference (for create, e.g., 'Sheet1!A1')")] string? value = null)
        => ExcelParameterTool.ExcelParameter(action, filePath, parameterName, value);

    // Cell Operations
    /// <summary>
    /// Manage individual Excel cells - values and formulas for precise control
    /// Delegates to ExcelCellTool for implementation.
    /// </summary>
    [McpServerTool(Name = "excel_cell")]
    [Description("Manage individual Excel cell values and formulas. Supports: get-value, set-value, get-formula, set-formula.")]
    public static string ExcelCell(
        [Description("Action: get-value, set-value, get-formula, set-formula")] string action,
        [Description("Excel file path (.xlsx or .xlsm)")] string filePath,
        [Description("Worksheet name")] string sheetName,
        [Description("Cell address (e.g., 'A1', 'B5')")] string cellAddress,
        [Description("Value or formula to set (for set-value/set-formula actions)")] string? value = null)
        => ExcelCellTool.ExcelCell(action, filePath, sheetName, cellAddress, value);

    // VBA Script Operations
    /// <summary>
    /// Manage Excel VBA scripts - modules, procedures, and macro execution (requires .xlsm files)
    /// Delegates to ExcelVbaTool for implementation.
    /// </summary>
    [McpServerTool(Name = "excel_vba")]
    [Description("Manage Excel VBA scripts and macros (requires .xlsm files). Supports: list, export, import, update, run, delete.")]
    public static string ExcelVba(
        [Description("Action: list, export, import, update, run, delete")] string action,
        [Description("Excel file path (must be .xlsm for VBA operations)")] string filePath,
        [Description("VBA module name or procedure name (format: 'Module.Procedure' for run)")] string? moduleName = null,
        [Description("VBA file path (.vba extension for import/export/update)")] string? vbaFilePath = null,
        [Description("Parameters for VBA procedure execution (comma-separated)")] string? parameters = null)
        => ExcelVbaTool.ExcelVba(action, filePath, moduleName, vbaFilePath, parameters);
}
