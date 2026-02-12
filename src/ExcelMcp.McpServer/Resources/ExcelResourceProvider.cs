using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Resources;

/// <summary>
/// MCP resources for documenting available Excel workbook URIs.
/// Resources help LLMs understand what can be inspected in Excel workbooks.
/// 
/// NOTE: MCP SDK 0.4.0-preview.2 does NOT support McpServerResourceTemplate yet.
/// Dynamic URI patterns (excel://{path}/queries/{name}) will be added when SDK supports it.
/// For now, use tools (powerquery list, etc.) for actual data retrieval.
/// </summary>
[McpServerResourceType]
public static class ExcelResourceProvider
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    /// <summary>
    /// Documents available Excel workbook resource URIs.
    /// </summary>
    [McpServerResource(UriTemplate = "excel://help/resources")]
    [Description("Guide to available Excel workbook resources")]
    public static Task<string> GetResourceGuide()
    {
        var guide = new
        {
            title = "Excel Workbook Resources",
            description = "URI patterns for inspecting Excel workbooks",
            note = "Use tools to retrieve actual data (MCP SDK resource templates not yet supported)",
            resourceTypes = new[]
            {
                new
                {
                    type = "Power Queries",
                    toolAction = "Use powerquery tool with action='list' to see all queries",
                    example = "powerquery(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "Worksheets",
                    toolAction = "Use worksheet tool with action='list' to see all worksheets",
                    example = "worksheet(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "Parameters (Named Ranges)",
                    toolAction = "Use namedrange tool with action='list' to see all parameters",
                    example = "namedrange(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "Data Model Tables",
                    toolAction = "Use datamodel tool with action='list-tables'",
                    example = "datamodel(action: 'list-tables', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "DAX Measures",
                    toolAction = "Use datamodel tool with action='list-measures'",
                    example = "datamodel(action: 'list-measures', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "VBA Modules",
                    toolAction = "Use vba tool with action='list'",
                    example = "vba(action: 'list', excelPath: 'workbook.xlsm')"
                },
                new
                {
                    type = "Excel Tables",
                    toolAction = "Use table tool with action='list'",
                    example = "table(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "Connections",
                    toolAction = "Use connection tool with action='list'",
                    example = "connection(action: 'list', excelPath: 'workbook.xlsx')"
                }
            },
            usage = new
            {
                discovery = "Use tool 'list' actions to discover workbook contents",
                inspection = "Use tool 'view' actions to examine specific items",
                modification = "Use other tool actions to create/update/delete items"
            },
            futureEnhancements = "Dynamic resource templates (excel://{path}/queries/{name}) will be added when MCP SDK supports McpServerResourceTemplate"
        };

        return Task.FromResult(JsonSerializer.Serialize(guide, JsonOptions));
    }

    /// <summary>
    /// Quick reference for common Excel operations.
    /// </summary>
    [McpServerResource(UriTemplate = "excel://help/quickref")]
    [Description("Quick reference for common Excel MCP operations")]
    public static Task<string> GetQuickReference()
    {
        var quickRef = new
        {
            title = "Excel MCP Quick Reference",
            commonOperations = new[]
            {
                new
                {
                    task = "List all Power Queries",
                    tool = "powerquery",
                    action = "list",
                    example = "powerquery(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    task = "View Power Query M code",
                    tool = "powerquery",
                    action = "view",
                    example = "powerquery(action: 'view', excelPath: 'workbook.xlsx', queryName: 'SalesData')"
                },
                new
                {
                    task = "Import query to Data Model",
                    tool = "powerquery",
                    action = "import",
                    example = "powerquery(action: 'import', excelPath: 'workbook.xlsx', queryName: 'Sales', sourcePath: 'sales.pq', loadDestination: 'data-model')"
                },
                new
                {
                    task = "List all worksheets",
                    tool = "worksheet",
                    action = "list",
                    example = "worksheet(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    task = "List all DAX measures",
                    tool = "datamodel",
                    action = "list-measures",
                    example = "datamodel(action: 'list-measures', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    task = "Get cell values",
                    tool = "range",
                    action = "get-values",
                    example = "range(action: 'get-values', excelPath: 'workbook.xlsx', sheetName: 'Data', rangeAddress: 'A1:D10')"
                },
                new
                {
                    task = "Work with sessions",
                    tool = "file",
                    action = "open/close",
                    example = "file(action: 'open') → operations with sessionId → file(action: 'close', save: true)"
                }
            },
            sessionWorkflow = new[]
            {
                "Open session: file(action: 'open', excelPath: '...')",
                "Use sessionId with all subsequent operations",
                "Close session: file(action: 'close', sessionId: '...', save: true)"
            }
        };

        return Task.FromResult(JsonSerializer.Serialize(quickRef, JsonOptions));
    }
}


