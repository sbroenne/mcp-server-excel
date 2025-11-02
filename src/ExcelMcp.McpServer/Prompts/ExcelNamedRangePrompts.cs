using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for Excel parameter (named range) management.
/// </summary>
[McpServerPromptType]
public static class ExcelNamedRangePrompts
{
    /// <summary>
    /// Guide for efficient parameter creation using bulk operations.
    /// </summary>
    [McpServerPrompt(Name = "excel_namedrange_bulk_guide")]
    [Description("Guide for creating multiple Excel parameters efficiently using bulk operations")]
    public static ChatMessage NamedRangeBulkGuide()
    {
        return new ChatMessage(ChatRole.User, @"When creating multiple Excel named range parameters, use the 'create-bulk' action for maximum efficiency.

# WHEN TO USE BULK PARAMETER CREATION

Use 'create-bulk' when:
- User asks to create 2+ parameters
- User mentions: ""parameters for..."", ""named ranges for...""
- You see a list of configuration values
- Setting up a new workbook with multiple inputs

# BULK vs INDIVIDUAL CREATION

❌ INEFFICIENT (10 calls for 5 parameters):
excel_namedrange(action: 'create', excelPath: 'file.xlsx', namedRangeName: 'Start_Date', value: 'Sheet1!A1')
excel_namedrange(action: 'set', excelPath: 'file.xlsx', parameterName: 'Start_Date', value: '2025-07-01')
excel_namedrange(action: 'create', excelPath: 'file.xlsx', namedRangeName: 'End_Date', value: 'Sheet1!A2')
excel_namedrange(action: 'set', excelPath: 'file.xlsx', parameterName: 'End_Date', value: '2025-12-31')
... (repeat 3 more times)

✅ EFFICIENT (1 call for 5 parameters):
excel_namedrange(action: 'create-bulk', excelPath: 'file.xlsx', parametersJson: JSON.stringify([
  { name: 'Start_Date', reference: 'Sheet1!$A$1', value: '2025-07-01' },
  { name: 'End_Date', reference: 'Sheet1!$A$2', value: '2025-12-31' },
  { name: 'Plan_Name', reference: 'Sheet1!$A$3', value: 'Q3 Plan' },
  { name: 'Duration_Months', reference: 'Sheet1!$A$4', value: 6 },
  { name: 'Region', reference: 'Sheet1!$A$5', value: 'West' }
]))

RESULT: 90% reduction in calls (10 → 1), single Excel session

# BEST PRACTICES

1. **Use absolute references**: 'Sheet1!$A$1' not 'Sheet1!A1'
2. **Set values immediately**: Include 'value' property for initial values
3. **Combine with batch mode**: For even more operations, wrap in begin_excel_batch
4. **JSON format**: Use lowercase property names (name, reference, value)

# COMMON PARAMETER PATTERNS

Configuration parameters:
- Date ranges: Start_Date, End_Date
- Periods: Fiscal_Year, Quarter, Month
- Filters: Region, Product_Category, Status
- Thresholds: Min_Value, Max_Value, Target
- Names: Project_Name, Department, Owner");
    }
}
