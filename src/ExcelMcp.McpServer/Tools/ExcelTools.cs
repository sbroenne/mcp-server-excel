#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' requirements

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel tools documentation and guidance for Model Context Protocol (MCP) server.
/// 
/// 🔧 Tool Architecture (6 Domain-Focused Tools):
/// - ExcelFileTool: File operations (create-empty)
/// - ExcelPowerQueryTool: M code and data loading management  
/// - ExcelWorksheetTool: Sheet operations and bulk data handling
/// - ExcelParameterTool: Named ranges as configuration parameters
/// - ExcelCellTool: Precise individual cell operations
/// - ExcelVbaTool: VBA macro management and execution
/// 
/// 🤖 LLM Usage Guidelines:
/// 1. Start with ExcelFileTool to create new Excel files
/// 2. Use ExcelWorksheetTool for data operations and sheet management
/// 3. Use ExcelPowerQueryTool for advanced data transformation
/// 4. Use ExcelParameterTool for configuration and reusable values
/// 5. Use ExcelCellTool for precision operations on individual cells
/// 6. Use ExcelVbaTool for complex automation (requires .xlsm files)
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
/// - ExcelCellTool.cs: excel_cell tool
/// - ExcelVbaTool.cs: excel_vba tool
/// 
/// This prevents duplicate tool registration conflicts in the MCP framework.
/// </summary>
public static class ExcelTools
{
    // This class now serves as documentation only.
    // All MCP tool registrations have been moved to individual tool files
    // to prevent duplicate registration conflicts with the MCP framework.
}
