using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// Quick reference prompt for Excel connection types and COM API limitations.
/// </summary>
[McpServerPromptType]
public static class ExcelConnectionPrompts
{
    /// <summary>
    /// Quick reference for Excel connection types and critical COM API limitations.
    /// </summary>
    [McpServerPrompt(Name = "excel_connection_reference")]
    [Description("Quick reference: Excel connection types, which ones work via COM API, and critical limitations")]
    public static ChatMessage ConnectionReference()
    {
        return new ChatMessage(ChatRole.User, @"# Excel Connection Types - Quick Reference

## 9 Connection Types in Excel

| Type | Name | COM API Status | Use For |
|------|------|----------------|---------|
| 1 | OLEDB | ❌ Connections.Add() FAILS | SQL Server, Access - **create via Excel UI or .odc files** |
| 2 | ODBC | ❌ Connections.Add() FAILS | ODBC sources - **create via Excel UI or .odc files** |
| 3 | TEXT | ✅ Works | CSV/text files - **recommended for automation** |
| 4 | WEB | ⚠️ Untested | Web queries |
| 5 | XMLMAP | ⚠️ Untested | XML data |
| 6 | DATAFEED | ⚠️ Untested | Data feeds |
| 7 | MODEL | ⚠️ Untested | Data model |
| 8 | WORKSHEET | ⚠️ Untested | Worksheet connections |
| 9 | NOSOURCE | ⚠️ Untested | No source |

## ⚠️ CRITICAL Limitation

**OLEDB/ODBC connections CANNOT be created via `Connections.Add()`** - Excel throws ""Value does not fall within expected range"". This is an Excel COM API limitation.

**User Workarounds:**
- Create OLEDB/ODBC in Excel UI (Data → Get Data → From Database)
- Import from .odc files (Office Data Connection files)
- Use excel_connection tool to **manage existing** connections (list, view, update, delete, refresh)

**For Automation:**
- Use TEXT connections (CSV files) - these work reliably via COM API
- Connection string format: `TEXT;C:\path\to\file.csv`

## Known Issue: Type 3/4 Confusion
Excel may return type 4 (WEB) when reading TEXT connections created with ""TEXT;filepath"" format. If user reports connection type mismatch, this is a known Excel behavior.");
    }
}
