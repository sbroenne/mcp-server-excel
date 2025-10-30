using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompts for teaching LLMs about Power Query and Data Model workflows.
/// Critical upfront knowledge to prevent common mistakes.
/// </summary>
[McpServerPromptType]
public static class ExcelPowerQueryDataModelPrompts
{
    [McpServerPrompt(Name = "excel_powerquery_datamodel_guide")]
    [Description("Essential guide: Understanding Power Query load destinations and Data Model workflows")]
    public static ChatMessage PowerQueryDataModelGuide()
    {
        return new ChatMessage(ChatRole.User, @"When working with Excel Power Query and Data Model (Power Pivot), understanding WHERE data loads is critical:

THREE LOAD DESTINATIONS:

1. WORKSHEET ONLY (set-load-to-table):
   - Data appears in worksheet as formatted table (users see it)
   - Data is NOT in Power Pivot Data Model
   - Cannot use DAX measures or relationships
   - Cannot add to Data Model using excel_table add-to-datamodel (will fail - no Excel Table object exists)

2. DATA MODEL ONLY (set-load-to-data-model):
   - Data loaded into Power Pivot (ready for DAX measures and relationships)
   - Data is NOT visible in any worksheet (connection-only to Data Model)
   - Use excel_datamodel tool for DAX measures, relationships, calculated columns

3. BOTH WORKSHEET AND DATA MODEL (set-load-to-both):
   - Data visible in worksheet AND available in Power Pivot
   - Best of both worlds: users see data, and you can create DAX measures
   - Use excel_datamodel tool for DAX operations

COMMON MISTAKE TO AVOID:

User says: 'Load this query to Data Model for DAX measures'

WRONG approach:
1. excel_powerquery action: set-load-to-table, targetSheet: 'Sales'
2. excel_table action: add-to-datamodel, tableName: 'Sales'  ← FAILS! No Excel Table exists
3. excel_table action: create, tableName: 'Sales', range: 'A1:Z100'  ← Workaround, but unnecessary
4. excel_table action: add-to-datamodel, tableName: 'Sales'  ← Finally works, but convoluted

RIGHT approach:
1. excel_powerquery action: set-load-to-data-model, queryName: 'Sales'  ← Done! Data in Power Pivot
2. excel_datamodel action: create-measure (or other DAX operations)

Or if user wants to SEE the data too:
1. excel_powerquery action: set-load-to-both, queryName: 'Sales', targetSheet: 'Sales'
2. excel_datamodel action: create-measure

WHEN TO USE EACH ACTION:

Use 'set-load-to-table' when:
- User wants to see data in Excel worksheet
- No DAX measures or Data Model needed
- Simple data viewing or manual analysis

Use 'set-load-to-data-model' when:
- User mentions: DAX, measures, relationships, Power Pivot, Data Model
- User wants analytical capabilities (measures, calculations across tables)
- Data doesn't need to be visible in worksheet

Use 'set-load-to-both' when:
- User wants BOTH visibility AND DAX capabilities
- User says: 'show the data and create measures'
- Best default for Data Model scenarios where users also want to see data

CHANGING LOAD DESTINATION:

If you already loaded to worksheet only and user NOW wants Data Model:
- Just call: excel_powerquery action: set-load-to-data-model
- No need to delete and recreate anything
- Power Query can change load destination anytime

WORKFLOW RESPONSES:

After set-load-to-table, you'll see:
- 'Query data loaded to worksheet (visible to users as formatted table)'
- 'IMPORTANT: This is NOT loaded to Power Pivot Data Model yet'
- This tells you: Data is visible but not available for DAX

After set-load-to-data-model, you'll see:
- 'Query data loaded to Power Pivot Data Model (ready for DAX)'
- 'Data is in model but NOT visible in worksheet'
- This tells you: Ready for DAX, but users won't see data in worksheet

After set-load-to-both, you'll see:
- 'Query data loaded to BOTH worksheet AND Power Pivot Data Model'
- 'Data visible in worksheet AND available for DAX measures/relationships'
- This tells you: Best of both worlds

REMEMBER: The load destination determines what you can DO with the data, not just where it appears!");
    }
}
