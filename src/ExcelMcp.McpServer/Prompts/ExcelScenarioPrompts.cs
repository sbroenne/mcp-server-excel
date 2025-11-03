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

## RECOMMENDED WORKFLOW (using batch mode for efficiency):

Step 1: begin_excel_batch(excelPath: 'FinancialReport.xlsx')
Step 2: excel_worksheet(action: 'create', sheetName: 'Report', batchId: '<batch-id>')
Step 3: excel_range(action: 'set-values', rangeAddress: 'A1:D1', values: [['Month', 'Revenue', 'Expenses', 'Profit']])
Step 4: excel_range(action: 'format-range', rangeAddress: 'A1:D1', bold: true, fillColor: '#4472C4')
Step 5: excel_range(action: 'set-formulas', rangeAddress: 'D2:D{monthCount + 1}', formulas: [['=B2-C2'], ...])
Step 6: excel_range(action: 'set-number-format', rangeAddress: 'B2:D{monthCount + 1}', formatCode: '$#,##0')
Step 7: commit_excel_batch(batchId: '<batch-id>', save: true)

RESULT: Professional formatted report with {monthCount} months of data
TIME SAVINGS: 95% faster with batch mode (2-3 seconds vs 20-25 seconds)
");
    }

    [McpServerPrompt(Name = "excel_multi_query_import")]
    [Description("Efficiently import multiple Power Queries to Data Model for DAX")]
    public static ChatMessage MultiQueryImport(
        [Description("Number of queries to import")] int? queryCount = null)
    {
        var count = queryCount ?? 4;

        return new ChatMessage(ChatRole.User, $@"
# IMPORT {count} POWER QUERIES TO DATA MODEL (EFFICIENT WORKFLOW)

Step 1: begin_excel_batch(excelPath: 'Analytics.xlsx')
Step 2-{count + 1}: excel_powerquery(action: 'import', queryName: '<name>', sourcePath: '<file>.pq', loadDestination: 'data-model', batchId: '<id>')
Step {count + 2}: commit_excel_batch(batchId: '<id>', save: true)

KEY: Use loadDestination: 'data-model' for direct Power Pivot loading
RESULT: {count} queries loaded and ready for DAX measures in 1-2 seconds
TIME SAVINGS: {count * 2} seconds → 1-2 seconds with batch mode
");
    }

    [McpServerPrompt(Name = "excel_build_data_entry_form")]
    [Description("Build a data entry form with dropdown validation and formatting")]
    public static ChatMessage BuildDataEntryForm()
    {
        return new ChatMessage(ChatRole.User, @"
# BUILD DATA ENTRY FORM WITH VALIDATION

Create professional data entry form with dropdowns, date validation, and formatted layout.

WORKFLOW:
1. begin_excel_batch + excel_worksheet(action: 'create', sheetName: 'Employee Form')
2. excel_range(action: 'set-values') - Add headers and labels
3. excel_range(action: 'format-range') - Professional formatting
4. excel_range(action: 'validate-range', validationType: 'list') - Department dropdown
5. excel_range(action: 'validate-range', validationType: 'date') - Hire date validation
6. excel_range(action: 'validate-range', validationType: 'list') - Status dropdown
7. commit_excel_batch(save: true)

RESULT: Professional form with validation, dropdowns, borders, and formatting
");
    }

    [McpServerPrompt(Name = "excel_version_control_workflow")]
    [Description("Workflow for exporting Excel code artifacts to Git version control")]
    public static ChatMessage VersionControlWorkflow()
    {
        return new ChatMessage(ChatRole.User, @"
# VERSION CONTROL WORKFLOW FOR EXCEL CODE

Export Power Query M code, VBA modules, and DAX measures to files for Git tracking.

EXPORT WORKFLOW:
1. excel_powerquery(action: 'export', queryName: '<name>', targetPath: 'queries/<name>.pq')
2. excel_vba(action: 'export', moduleName: '<name>', targetPath: 'vba/<name>.bas')
3. excel_datamodel(action: 'export-measure', targetPath: 'dax/measures.dax')

GIT WORKFLOW:
git add queries/*.pq vba/*.bas dax/*.dax
git commit -m 'Export Excel code artifacts'
git push origin main

IMPORT BACK:
excel_powerquery(action: 'import', sourcePath: 'queries/<name>.pq', loadDestination: 'data-model')
excel_vba(action: 'import', sourcePath: 'vba/<name>.bas')

BENEFITS: Track changes, code review, rollback, collaboration, audit trail
");
    }

    [McpServerPrompt(Name = "excel_build_analytics_workbook")]
    [Description("Complete workflow: Build analytics workbook with Power Query, Data Model, DAX measures")]
    public static ChatMessage BuildAnalyticsWorkbook()
    {
        return new ChatMessage(ChatRole.User, @"
# BUILD COMPLETE ANALYTICS WORKBOOK

End-to-end: Import data → Build Data Model → Create DAX measures → Add PivotTable

WORKFLOW (using batch mode):
1. begin_excel_batch(excelPath: 'Analytics.xlsx')
2. Import 4 queries with loadDestination: 'data-model' (Sales, Products, Customers, Calendar)
3. Create 3 relationships (Sales→Products, Sales→Customers, Sales→Calendar)
4. Create 4 DAX measures (Total Revenue, Revenue vs Budget %, Avg Order Value, Customer Count)
5. commit_excel_batch(save: true)
6. excel_pivottable(action: 'create', sourceType: 'data-model')
7. Add fields to PivotTable (rows, columns, values)

RESULT: 4 data sources, 3 relationships, 4 DAX measures, 1 PivotTable
TIME: 3-5 seconds with batch mode (vs 30-40 seconds without)
");
    }
}
