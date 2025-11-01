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

THREE LOAD DESTINATIONS (use loadDestination parameter on import action):

1. WORKSHEET ONLY (loadDestination: 'worksheet'):
   - Data appears in worksheet as formatted table (users see it)
   - Data is NOT in Power Pivot Data Model
   - Cannot use DAX measures or relationships
   - Use for simple data viewing without analytics

2. DATA MODEL ONLY (loadDestination: 'data-model'):
   - Data loaded into Power Pivot (ready for DAX measures and relationships)
   - Data is NOT visible in any worksheet (connection-only to Data Model)
   - Use excel_datamodel tool for DAX measures, relationships, calculated columns
   - BEST for analytical workflows with DAX

3. BOTH WORKSHEET AND DATA MODEL (loadDestination: 'both'):
   - Data visible in worksheet AND available in Power Pivot
   - Best of both worlds: users see data, and you can create DAX measures
   - Use excel_datamodel tool for DAX operations

4. CONNECTION-ONLY (loadDestination: 'connection-only'):
   - M code imported but NOT executed
   - No data loading or validation
   - Advanced use only

RECOMMENDED WORKFLOW - Load to Data Model in ONE CALL:

User says: 'Import these 4 queries to Data Model for DAX measures'

✅ RIGHT approach (using loadDestination parameter):
excel_powerquery(action: Import, queryName: 'Sales', sourcePath: 'sales.pq', loadDestination: 'data-model')
excel_powerquery(action: Import, queryName: 'Products', sourcePath: 'products.pq', loadDestination: 'data-model')
excel_powerquery(action: Import, queryName: 'Customers', sourcePath: 'customers.pq', loadDestination: 'data-model')
excel_powerquery(action: Import, queryName: 'Regions', sourcePath: 'regions.pq', loadDestination: 'data-model')
→ 4 calls total, data ready for DAX immediately

❌ DEPRECATED approach (OLD - don't use):
excel_powerquery(action: Import, queryName: 'Sales', sourcePath: 'sales.pq')
excel_powerquery(action: SetLoadToDataModel, queryName: 'Sales')
... (repeat for 3 more queries)
→ 8 calls total - INEFFICIENT!

WHEN TO USE EACH DESTINATION:

Use 'worksheet' (default) when:
- User wants to see data in Excel worksheet
- No DAX measures or Data Model needed
- Simple data viewing or manual analysis

Use 'data-model' when:
- User mentions: DAX, measures, relationships, Power Pivot, Data Model
- User wants analytical capabilities (measures, calculations across tables)
- Data doesn't need to be visible in worksheet

Use 'both' when:
- User wants BOTH visibility AND DAX capabilities
- User says: 'show the data and create measures'
- Best default for Data Model scenarios where users also want to see data

Use 'connection-only' when:
- Advanced scenarios: M code import without validation
- Rarely needed

CHANGING LOAD DESTINATION:

If you already loaded to worksheet and user NOW wants Data Model:
- excel_powerquery(action: SetLoadToDataModel, queryName: 'Sales')
- No need to delete and recreate anything
- Power Query can change load destination anytime

REFRESH WITH LOAD DESTINATION:

If query is connection-only and user wants to refresh AND load data:
- excel_powerquery(action: Refresh, queryName: 'Sales', loadDestination: 'worksheet')
- ONE call instead of two (set-load + refresh)
- Applies load configuration then refreshes data
- Also works with: loadDestination: 'data-model' or 'both'

WORKFLOW RESPONSES (what to expect):

After loadDestination: 'worksheet', you'll see:
- 'Query data loaded to worksheet (visible to users as formatted table)'
- This tells you: Data is visible but not available for DAX

After loadDestination: 'data-model', you'll see:
- 'Query data loaded to Power Pivot Data Model (ready for DAX)'
- This tells you: Ready for DAX, but users won't see data in worksheet

After loadDestination: 'both', you'll see:
- 'Query data loaded to BOTH worksheet AND Power Pivot Data Model'
- This tells you: Best of both worlds

EXCEL_TABLE VS EXCEL_POWERQUERY - WHEN TO USE EACH:

excel_powerquery tool:
- For EXTERNAL data sources (databases, web APIs, files, SharePoint, etc.)
- Loads data FROM outside Excel INTO Excel
- Creates Power Query connections with M code
- Examples: Load sales data from SQL Server, import CSV files, pull data from web APIs
- Actions: import (with loadDestination!), refresh, set-load-to-data-model, etc.

excel_table tool:
- For data ALREADY in Excel worksheets (ranges of cells)
- Converts existing ranges to Excel Tables (ListObject)
- Provides structure: AutoFilter, structured references ([@Column]), dynamic expansion
- Examples: Convert a range A1:Z100 to a table, add AutoFilter
- Actions: create, resize, add-column, apply-filter, sort, add-to-datamodel

Use excel_powerquery when:
- Data comes from EXTERNAL sources (not already in Excel)
- You need to refresh data from source periodically
- You want Power Query M code transformations

Use excel_table when:
- Data is ALREADY in Excel worksheet as a range
- You want to add structure (AutoFilter, formulas with [@Column])
- You have manually entered data or pasted data

REMEMBER: Use the loadDestination parameter on import for efficient Data Model workflows!");
    }
}
