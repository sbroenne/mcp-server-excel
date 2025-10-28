#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' requirements

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel tools documentation and guidance for Model Context Protocol (MCP) server.
///
/// 🔧 Tool Architecture (8 Domain-Focused Tools):
/// - ExcelFileTool: File operations (create-empty)
/// - ExcelPowerQueryTool: M code and data loading management
/// - ExcelWorksheetTool: Sheet lifecycle management (create, rename, copy, delete)
/// - ExcelParameterTool: Named ranges as configuration parameters
/// - ExcelRangeTool: Unified range operations (values, formulas, formatting, hyperlinks)
/// - ExcelVbaTool: VBA macro management and execution
/// - ExcelDataModelTool: Data Model (Power Pivot) operations
/// - ExcelVersionTool: Check for updates on NuGet.org
///
/// 🤖 LLM Usage Guidelines:
/// 1. Start with ExcelFileTool to create new Excel files
/// 2. Use ExcelWorksheetTool for sheet lifecycle (create, rename, copy, delete)
/// 3. Use ExcelRangeTool for ALL data operations (read, write, formulas, formatting, hyperlinks)
/// 4. Use ExcelPowerQueryTool for advanced data transformation
/// 5. Use ExcelParameterTool for configuration and reusable values
/// 6. Use ExcelVbaTool for complex automation (requires .xlsm files)
/// 7. Use ExcelDataModelTool for Data Model and DAX operations
/// 8. Use ExcelVersionTool to check for updates
///
/// 📝 Parameter Patterns:
/// - action: Always the first parameter, defines what operation to perform
/// - filePath/excelPath: Excel file path (.xlsx or .xlsm based on requirements)
/// - Context-specific parameters: Each tool has domain-appropriate parameters
///
/// 🎯 Design Philosophy:
/// - Resource-based: Tools represent Excel domains, not individual operations
/// - Action-oriented: Each tool supports multiple related actions
/// - LLM-friendly: Clear naming, comprehensive documentation, predictable patterns
/// - Error-consistent: Standardized error handling across all tools
///
/// 🚨 IMPORTANT: This class NO LONGER contains MCP tool registrations!
/// All tools are now registered individually in their respective classes with [McpServerToolType]:
/// - ExcelFileTool.cs: excel_file tool
/// - ExcelPowerQueryTool.cs: excel_powerquery tool
/// - ExcelWorksheetTool.cs: excel_worksheet tool
/// - ExcelParameterTool.cs: excel_parameter tool
/// - ExcelRangeTool.cs: excel_range tool (replaces excel_cell)
/// - ExcelVbaTool.cs: excel_vba tool
/// - ExcelDataModelTool.cs: excel_datamodel tool
/// - ExcelVersionTool.cs: excel_version tool
///
/// This prevents duplicate tool registration conflicts in the MCP framework.
/// </summary>
public static class ExcelTools
{
    // This class now serves as documentation only.
    // All MCP tool registrations have been moved to individual tool files
    // to prevent duplicate registration conflicts with the MCP framework.
}
