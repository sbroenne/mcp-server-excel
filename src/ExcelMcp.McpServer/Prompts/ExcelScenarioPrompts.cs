using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for common Excel workflow scenarios.
/// Provides step-by-step templates for frequent tasks.
/// </summary>
[McpServerPromptType]
public static class ExcelScenarioPrompts
{
    [McpServerPrompt(Name = "excel_build_financial_report")]
    [Description("Step-by-step guide to build a formatted financial report with formulas")]
    public static ChatMessage BuildFinancialReport(
        [Description("Report title (optional)")] string? reportTitle = null,
        [Description("Number of months (default: 12)")] int? months = null)
    {
        var monthCount = months ?? 12;
        var title = reportTitle ?? "Monthly Revenue Report";

        return new ChatMessage(ChatRole.User, $@"
# BUILD FINANCIAL REPORT: {title}

Complete workflow for creating a professional financial report with formulas and formatting.

**CRITICAL: Keep session open until ALL steps complete - do NOT close prematurely**

## RECOMMENDED WORKFLOW:

1. excel_file(action: 'open', excelPath: 'FinancialReport.xlsx')
   \u2192 Returns sessionId (use for ALL remaining operations)
2. excel_worksheet(action: 'create', sheetName: 'Report', sessionId: '<sessionId>')
3. excel_range(action: 'set-values', rangeAddress: 'A1:D1', values: [['Month', 'Revenue', 'Expenses', 'Profit']], sessionId: '<sessionId>')
4. excel_range(action: 'format-range', rangeAddress: 'A1:D1', bold: true, fillColor: '#4472C4', sessionId: '<sessionId>')
5. excel_range(action: 'set-formulas', rangeAddress: 'D2:D{monthCount + 1}', formulas: [['=B2-C2'], ...], sessionId: '<sessionId>')
6. excel_range(action: 'set-number-format', rangeAddress: 'B2:D{monthCount + 1}', formatCode: '$#,##0', sessionId: '<sessionId>')
7. excel_file(action: 'close', save: true, sessionId: '<sessionId>')
   \u2192 ONLY close when report is complete

RESULT: Professional formatted report with {monthCount} months of data
");
    }

    [McpServerPrompt(Name = "excel_multi_query_import")]
    [Description("Efficiently import multiple Power Queries to Data Model for DAX")]
    public static ChatMessage MultiQueryImport(
        [Description("Number of queries to import")] int? queryCount = null)
    {
        var count = queryCount ?? 4;

        return new ChatMessage(ChatRole.User, $@"
# IMPORT {count} POWER QUERIES TO DATA MODEL

Complete workflow for importing multiple queries and preparing for DAX analysis.

**CRITICAL: Keep session open across ALL query imports - do NOT close between operations**

## RECOMMENDED WORKFLOW:

1. excel_file(action: 'open', excelPath: 'Analytics.xlsx')
   \u2192 Returns sessionId (use for ALL remaining operations)
2. For each query:
   - excel_powerquery(action: 'create', queryName: '<name>', mCode: '<M code>', loadDestination: 'data-model', sessionId: '<sessionId>')
3. excel_file(action: 'close', save: true, sessionId: '<sessionId>')
   \u2192 ONLY close after ALL queries imported

KEY: Use loadDestination: 'data-model' for direct Power Pivot loading
RESULT: {count} queries loaded and ready for DAX measures
");
    }

    [McpServerPrompt(Name = "excel_build_data_entry_form")]
    [Description("Build a data entry form with dropdown validation and formatting")]
    public static ChatMessage BuildDataEntryForm()
    {
        return new ChatMessage(ChatRole.User, @"
# BUILD DATA ENTRY FORM WITH VALIDATION

Create professional data entry form with dropdowns, date validation, and formatted layout.

**CRITICAL: Keep session open until form is complete - do NOT close between operations**

WORKFLOW:
1. excel_file(action: 'open', excelPath: 'DataEntryForm.xlsx')
   \u2192 Returns sessionId (use for ALL remaining operations)
2. excel_worksheet(action: 'create', sheetName: 'Employee Form', sessionId: '<sessionId>')
3. excel_range(action: 'set-values', values: [['Employee ID', 'Name', 'Department', 'Status', 'Hire Date']], sessionId: '<sessionId>')
4. excel_range(action: 'format-range', rangeAddress: 'A1:E1', bold: true, fillColor: '#D9E1F2', sessionId: '<sessionId>')
5. excel_range(action: 'validate-range', rangeAddress: 'C2:C100', validationType: 'list', validationFormula1: 'IT,HR,Finance,Operations', sessionId: '<sessionId>')
6. excel_range(action: 'validate-range', rangeAddress: 'D2:D100', validationType: 'list', validationFormula1: 'Active,Inactive,Leave', sessionId: '<sessionId>')
7. excel_range(action: 'validate-range', rangeAddress: 'E2:E100', validationType: 'date', sessionId: '<sessionId>')
8. excel_file(action: 'close', save: true, sessionId: '<sessionId>')
   \u2192 ONLY close when form is complete

RESULT: Professional form with validation, dropdowns, and formatting
");
    }


    [McpServerPrompt(Name = "excel_build_analytics_workbook")]
    [Description("Complete workflow: Build analytics workbook with Power Query, Data Model, DAX measures")]
    public static ChatMessage BuildAnalyticsWorkbook()
    {
        return new ChatMessage(ChatRole.User, @"
# BUILD COMPLETE ANALYTICS WORKBOOK

End-to-end: Import data → Build Data Model → Create DAX measures → Add PivotTable

**CRITICAL: Keep session open across ALL steps - do NOT close between operations**

WORKFLOW:
1. excel_file(action: 'open', excelPath: 'Analytics.xlsx')
   → Returns sessionId (use for ALL remaining operations)
2. Import 4 queries with loadDestination: 'data-model' (Sales, Products, Customers, Calendar)
   - excel_powerquery(action: 'create', queryName: 'Sales', mCode: '<M code>', loadDestination: 'data-model', sessionId: '<sessionId>')
   - ... (repeat for Products, Customers, Calendar)
3. Create 3 relationships
   - excel_datamodel(action: 'create-relationship', fromTable: 'Sales', fromColumn: 'ProductID', toTable: 'Products', toColumn: 'ProductID', sessionId: '<sessionId>')
   - excel_datamodel(action: 'create-relationship', fromTable: 'Sales', fromColumn: 'CustomerID', toTable: 'Customers', toColumn: 'CustomerID', sessionId: '<sessionId>')
   - excel_datamodel(action: 'create-relationship', fromTable: 'Sales', fromColumn: 'DateID', toTable: 'Calendar', toColumn: 'DateID', sessionId: '<sessionId>')
4. Create 4 DAX measures
   - excel_datamodel(action: 'create-measure', tableName: 'Measures', measureName: 'Total Revenue', daxFormula: 'SUM(Sales[Amount])', sessionId: '<sessionId>')
   - ... (repeat for other measures)
5. excel_pivottable(action: 'create-from-datamodel', dataModelTableName: 'Sales', destinationSheet: 'PivotTable', destinationCell: 'A1', sessionId: '<sessionId>')
6. excel_file(action: 'close', save: true, sessionId: '<sessionId>')
   → ONLY close when analytics workbook is complete

RESULT: 4 data sources, 3 relationships, 4 DAX measures, 1 PivotTable
");
    }
}
