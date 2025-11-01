using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Resources;

/// <summary>
/// MCP resources for documenting available Excel workbook URIs.
/// Resources help LLMs understand what can be inspected in Excel workbooks.
/// Note: Actual data retrieval should use tools (excel_powerquery list, etc.)
/// These resources serve as URI documentation and discovery aids.
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
            resourceTypes = new[]
            {
                new
                {
                    type = "Power Queries",
                    uriPattern = "Use excel_powerquery tool with action='list' to see all queries",
                    example = "excel_powerquery(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "Worksheets",
                    uriPattern = "Use excel_worksheet tool with action='list' to see all worksheets",
                    example = "excel_worksheet(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "Parameters (Named Ranges)",
                    uriPattern = "Use excel_parameter tool with action='list' to see all parameters",
                    example = "excel_parameter(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "Data Model Tables",
                    uriPattern = "Use excel_datamodel tool with action='list-tables'",
                    example = "excel_datamodel(action: 'list-tables', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "DAX Measures",
                    uriPattern = "Use excel_datamodel tool with action='list-measures'",
                    example = "excel_datamodel(action: 'list-measures', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "VBA Modules",
                    uriPattern = "Use excel_vba tool with action='list'",
                    example = "excel_vba(action: 'list', excelPath: 'workbook.xlsm')"
                },
                new
                {
                    type = "Excel Tables",
                    uriPattern = "Use excel_table tool with action='list'",
                    example = "excel_table(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    type = "Connections",
                    uriPattern = "Use excel_connection tool with action='list'",
                    example = "excel_connection(action: 'list', excelPath: 'workbook.xlsx')"
                }
            },
            usage = new
            {
                discovery = "Use tool 'list' actions to discover workbook contents",
                inspection = "Use tool 'view' actions to examine specific items",
                modification = "Use other tool actions to create/update/delete items"
            }
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
                    tool = "excel_powerquery",
                    action = "list",
                    example = "excel_powerquery(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    task = "View Power Query M code",
                    tool = "excel_powerquery",
                    action = "view",
                    example = "excel_powerquery(action: 'view', excelPath: 'workbook.xlsx', queryName: 'SalesData')"
                },
                new
                {
                    task = "Import query to Data Model",
                    tool = "excel_powerquery",
                    action = "import",
                    example = "excel_powerquery(action: 'import', excelPath: 'workbook.xlsx', queryName: 'Sales', sourcePath: 'sales.pq', loadDestination: 'data-model')"
                },
                new
                {
                    task = "List all worksheets",
                    tool = "excel_worksheet",
                    action = "list",
                    example = "excel_worksheet(action: 'list', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    task = "List all DAX measures",
                    tool = "excel_datamodel",
                    action = "list-measures",
                    example = "excel_datamodel(action: 'list-measures', excelPath: 'workbook.xlsx')"
                },
                new
                {
                    task = "Get cell values",
                    tool = "excel_range",
                    action = "get-values",
                    example = "excel_range(action: 'get-values', excelPath: 'workbook.xlsx', sheetName: 'Data', rangeAddress: 'A1:D10')"
                },
                new
                {
                    task = "Create multiple items efficiently",
                    tool = "begin_excel_batch",
                    action = "batch mode workflow",
                    example = "begin_excel_batch → multiple operations → commit_excel_batch"
                }
            },
            batchModeKeywords = new[]
            {
                "Use batch mode when you see: numbers (2+, 3+, 4+), plurals (queries, measures, worksheets), lists",
                "Pattern: begin_excel_batch → operations with batchId → commit_excel_batch(save: true)",
                "Benefit: 75-95% faster (single Excel session for all operations)"
            }
        };

        return Task.FromResult(JsonSerializer.Serialize(quickRef, JsonOptions));
    }
}
