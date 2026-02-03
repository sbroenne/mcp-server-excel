#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' requirements

namespace Sbroenne.ExcelMcp.McpServer.Tools;

/// <summary>
/// Excel tools documentation and guidance for Model Context Protocol (MCP) server.
///
/// Parameter Patterns:
/// - action: Always the first parameter, defines what operation to perform
/// - filePath/excelPath: Excel file path (.xlsx or .xlsm based on requirements)
/// - Context-specific parameters: Each tool has domain-appropriate parameters
///
/// Design Philosophy:
/// - Resource-based: Tools represent Excel domains, not individual operations
/// - Action-oriented: Each tool supports multiple related actions
/// - LLM-friendly: Clear naming, comprehensive documentation, predictable patterns
/// - Error-consistent: Standardized error handling across all tools
///
/// 🚨 IMPORTANT: This class NO LONGER contains MCP tool registrations!
/// All tools are now registered individually in their respective classes with [McpServerToolType]:
///
/// This prevents duplicate tool registration conflicts in the MCP framework.
/// </summary>
public static class ExcelTools
{
 // This class now serves as documentation only.
 // All MCP tool registrations have been moved to individual tool files
 // to prevent duplicate registration conflicts with the MCP framework.
}
