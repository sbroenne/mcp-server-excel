using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for gathering required information before executing Excel operations.
/// Acts as pre-flight checklists to prevent back-and-forth with users.
/// </summary>
[McpServerPromptType]
public static class ExcelElicitationPrompts
{
    [McpServerPrompt(Name = "excel_powerquery_checklist")]
    [Description("Checklist of information needed before importing a Power Query")]
    public static ChatMessage PowerQueryChecklist()
    {
        return new ChatMessage(ChatRole.User, @"
# BEFORE IMPORTING POWER QUERY - GATHER THIS INFO

REQUIRED:
☐ Query name (what to call it in Excel)
☐ Source file path (.pq file location)
☐ Excel file path (destination workbook)

RECOMMENDED (avoid second call):
☐ Load destination:
  - 'worksheet' (default - users see data in Excel)
  - 'data-model' (for DAX measures and Power Pivot)
  - 'both' (visible in worksheet AND available for DAX)
  - 'connection-only' (advanced - M code imported but not executed)

OPTIONAL:
☐ Target sheet name (if loadDestination: 'worksheet' or 'both')
☐ Privacy level (None, Private, Organizational, Public)
☐ Batch mode? (if importing 2+ queries, START with begin_excel_batch)

ASK USER FOR MISSING INFO before calling excel_powerquery.
BATCH MODE: Detect keywords (numbers, plurals, lists) → use begin_excel_batch automatically
");
    }

    [McpServerPrompt(Name = "excel_dax_measure_checklist")]
    [Description("Information needed before creating DAX measures")]
    public static ChatMessage DaxMeasureChecklist()
    {
        return new ChatMessage(ChatRole.User, @"
# BEFORE CREATING DAX MEASURES - GATHER THIS INFO

REQUIRED FOR EACH MEASURE:
☐ Measure name (e.g., 'Total Sales', 'Avg Price', 'Customer Count')
☐ Target table (which Data Model table owns this measure)
☐ DAX formula (e.g., 'SUM(Sales[Amount])', 'AVERAGE(Sales[Price])')

RECOMMENDED:
☐ Format string:
  - '#,##0.00' for decimals
  - '$#,##0' for currency
  - '0.00%' for percentage
  - 'General Number' for general
☐ Display folder (organize measures in categories like 'Revenue', 'Orders')
☐ Description (helps other users understand purpose)

WORKFLOW OPTIMIZATION:
☐ Are you creating 2+ measures? → Use batch mode (begin_excel_batch)
☐ Do Data Model tables exist? → Check with excel_datamodel(action: 'list-tables') first
☐ Is data loaded to Data Model? → Queries must use loadDestination: 'data-model' or 'both'

ASK USER FOR MISSING INFO.
BATCH MODE saves 75-95% time for multiple measures.
");
    }

    [McpServerPrompt(Name = "excel_range_formatting_checklist")]
    [Description("Information needed before formatting Excel ranges")]
    public static ChatMessage RangeFormattingChecklist()
    {
        return new ChatMessage(ChatRole.User, @"
# BEFORE FORMATTING RANGES - GATHER THIS INFO

REQUIRED:
☐ Excel file path
☐ Worksheet name
☐ Range address (e.g., 'A1:D10', 'SalesData')

FORMATTING OPTIONS (choose what applies):

FONT:
☐ Font name (Arial, Calibri, Times New Roman)
☐ Font size (8, 10, 11, 12, 14, 16)
☐ Bold, italic, underline
☐ Font color (#RRGGBB hex code, e.g., #FF0000 for red)

FILL:
☐ Fill color (#RRGGBB hex code, e.g., #FFFF00 for yellow)

BORDERS:
☐ Border style (none, continuous, dash, dot, double)
☐ Border weight (hairline, thin, medium, thick)
☐ Border color (#RRGGBB hex code)

ALIGNMENT:
☐ Horizontal (left, center, right, justify)
☐ Vertical (top, center, bottom)
☐ Wrap text (true/false)

NUMBER FORMAT:
☐ Format code (e.g., '$#,##0.00', '0.00%', 'mm/dd/yyyy')

COMMON PATTERNS:
- Headers: bold, fillColor: '#4472C4', fontColor: '#FFFFFF', center aligned
- Currency: formatCode: '$#,##0.00', right aligned
- Dates: formatCode: 'mm/dd/yyyy'
- Percentages: formatCode: '0.00%'

ASK USER what formatting they want before calling excel_range(action: 'format-range')
");
    }

    [McpServerPrompt(Name = "excel_data_validation_checklist")]
    [Description("Information needed before adding data validation to ranges")]
    public static ChatMessage DataValidationChecklist()
    {
        return new ChatMessage(ChatRole.User, @"
# BEFORE ADDING DATA VALIDATION - GATHER THIS INFO

REQUIRED:
☐ Excel file path
☐ Worksheet name
☐ Range address (cells to validate)
☐ Validation type:
  - 'list' (dropdown with fixed choices)
  - 'decimal' (number with min/max)
  - 'whole' (integer with min/max)
  - 'date' (date range)
  - 'custom' (formula-based validation)

TYPE-SPECIFIC INFO:

FOR LIST VALIDATION:
☐ List values (comma-separated, e.g., 'Active,Inactive,Pending')
☐ Show dropdown? (true recommended)

FOR NUMBER VALIDATION (decimal/whole):
☐ Operator (between, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, equal, notEqual)
☐ Minimum value
☐ Maximum value (if 'between' operator)

FOR DATE VALIDATION:
☐ Operator (same as number)
☐ Start date or formula (e.g., '2025-01-01' or '=TODAY()')
☐ End date (if 'between' operator)

FOR CUSTOM VALIDATION:
☐ Formula (e.g., '=COUNTIF($A$2:$A$100,A2)=1' for unique values)

OPTIONAL (recommended):
☐ Error style (stop, warning, information)
☐ Error title (short heading)
☐ Error message (helpful description)
☐ Input message (hint shown when cell selected)

ASK USER for validation requirements before calling excel_range(action: 'validate-range')
");
    }

    [McpServerPrompt(Name = "excel_batch_mode_detection")]
    [Description("Checklist for detecting when to use batch mode")]
    public static ChatMessage BatchModeDetection()
    {
        return new ChatMessage(ChatRole.User, @"
# DETECT BATCH MODE OPPORTUNITIES

ALWAYS use batch mode when you see ANY of these keywords in user request:

NUMBERS/COUNTS:
☐ '2 queries', '3 measures', '4 files', '5 parameters'
☐ 'multiple', 'several', 'a few', 'many'

PLURAL WORDS:
☐ 'queries', 'measures', 'relationships', 'parameters', 'worksheets', 'ranges'

LISTS:
☐ User provides list of items (sales.pq, products.pq, customers.pq)
☐ Enumerated items ('First import X, then Y, then Z')

REPETITIVE WORDS:
☐ 'each', 'all', 'every', 'for each'

WORKFLOW PATTERNS:
☐ 'set up a workbook with...' (usually multiple operations)
☐ 'build a report with...' (usually multiple operations)
☐ 'create a dashboard with...' (usually multiple operations)

DECISION TREE:
1. Count operations needed (look for keywords above)
2. If 2+ operations on SAME file → begin_excel_batch FIRST
3. If 1 operation only → Call tool directly (no batch)
4. After all operations → commit_excel_batch(save: true)

EXAMPLES REQUIRING BATCH MODE:
✅ 'Import these 4 queries to Data Model'
✅ 'Create parameters for Start_Date, End_Date, Region'
✅ 'Set up worksheets for Sales, Products, Customers'
✅ 'Build a report with headers, data, and formulas'

EXAMPLES NOT REQUIRING BATCH MODE:
❌ 'Just import this one query'
❌ 'Export the SalesData query'
❌ 'List all Power Queries'

REMEMBER: Batch mode detection should happen UPFRONT, not after seeing warnings!
");
    }

    [McpServerPrompt(Name = "excel_troubleshooting_guide")]
    [Description("Common error patterns and how to fix them")]
    public static ChatMessage TroubleshootingGuide()
    {
        return new ChatMessage(ChatRole.User, @"
# EXCEL MCP TROUBLESHOOTING GUIDE

COMMON ERRORS & FIXES:

ERROR: 'Power Query not found'
FIX: Check spelling. Use excel_powerquery(action: 'list') to see actual query names

ERROR: 'Worksheet not found'
FIX: Use excel_worksheet(action: 'list') to see available worksheets

ERROR: 'Named range not found'
FIX: Use excel_namedrange(action: 'list') to see all parameters

ERROR: 'Table not in Data Model'
FIX: Query must be loaded with loadDestination: 'data-model' or 'both'
     Use excel_powerquery(action: 'set-load-to-data-model', queryName: '<name>')

ERROR: 'VBA not supported in .xlsx file'
FIX: VBA requires .xlsm extension. Change file extension or create new .xlsm file

ERROR: 'Range address invalid'
FIX: Use format 'A1:Z100' (column letter + row number), NOT 'A1-Z100'

ERROR: 'Batch not found'
FIX: Call begin_excel_batch BEFORE using batchId parameter

ERROR: 'Multiple operations slow'
FIX: Detect keywords (numbers, plurals) → use batch mode automatically

ERROR: 'Cannot create relationship - column not found'
FIX: Check exact column names with excel_datamodel(action: 'list-columns', tableName: '<table>')

ERROR: 'DAX formula invalid'
FIX: Validate syntax. Common mistakes:
     - Missing quotes around text values
     - Wrong table/column names
     - Missing square brackets [Column]

ERROR: 'Privacy level error when combining sources'
FIX: Set privacyLevel parameter on import (None, Private, Organizational, Public)

PREVENTION:
- Always list items first (queries, worksheets, measures) before operating on them
- Use batch mode for 2+ operations
- Check error messages for specific details
- Validate names and paths before calling tools
");
    }
}
