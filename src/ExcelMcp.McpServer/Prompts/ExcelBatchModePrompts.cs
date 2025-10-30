using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// Essential guide for using Excel batch mode to achieve 95% faster operations.
/// </summary>
[McpServerPromptType]
public static class ExcelBatchModePrompts
{
    /// <summary>
    /// Essential guide: When and how to use Excel batch mode for 95% faster operations.
    /// </summary>
    [McpServerPrompt(Name = "excel_batch_mode_guide")]
    [Description("Essential guide: When and how to use Excel batch mode for 75-95% faster operations")]
    public static ChatMessage BatchModeGuide()
    {
        return new ChatMessage(ChatRole.User, @"When working with Excel files, you have access to a batch mode that makes multiple operations 75-95% faster (2-5 seconds per operation → <100ms per operation).

# CRITICAL: When to Use Batch Mode (AUTO-DETECT KEYWORDS)

**ALWAYS use batch mode when you see ANY of these in the user's request:**

KEYWORD TRIGGERS FOR BATCH MODE:
- Numbers: 'import these 4 queries', 'create 5 parameters', '3 measures', 'multiple'
- Plural words: 'queries', 'parameters', 'measures', 'relationships', 'tables', 'worksheets'
- Lists: User provides a list of items to process
- Repetitive words: 'each', 'all', 'every'

SPECIFIC EXAMPLES THAT REQUIRE BATCH MODE:
✅ 'Import these 4 .pq files into the Data Model' → 4 imports = BATCH MODE
✅ 'Create parameters for Start_Date, End_Date, Region, Product, Customer' → 5 parameters = BATCH MODE
✅ 'Create DAX measures for TotalSales, AvgPrice, ProductCount' → 3 measures = BATCH MODE
✅ 'Set up a workbook with Sales, Products, and Customers worksheets' → 3 worksheets = BATCH MODE
✅ 'Load all queries to data model' → Multiple queries = BATCH MODE
❌ 'Just import this one query' → Single operation = NO BATCH MODE

# The Three-Step Pattern (MANDATORY for 2+ operations)

**STEP 1:** FIRST call begin_excel_batch with the filePath
**STEP 2:** Pass the returned batchId to ALL subsequent tool calls on that file
**STEP 3:** FINALLY call commit_excel_batch with save: true

# Complete Example: Import 4 Power Queries to Data Model

**WRONG WAY (what LLMs naturally do - 8 calls, 8 Excel sessions, 16-24 seconds):**
excel_powerquery(action: 'import', excelPath: 'workbook.xlsx', queryName: 'Sales', sourcePath: 'sales.pq', loadToWorksheet: false)
excel_powerquery(action: 'set-load-to-data-model', excelPath: 'workbook.xlsx', queryName: 'Sales')
excel_powerquery(action: 'import', excelPath: 'workbook.xlsx', queryName: 'Products', sourcePath: 'products.pq', loadToWorksheet: false)
excel_powerquery(action: 'set-load-to-data-model', excelPath: 'workbook.xlsx', queryName: 'Products')
... (4 more calls for 2 more queries)

**RIGHT WAY (batch mode - 6 calls, 1 Excel session, 3-4 seconds):**
batch = begin_excel_batch(excelPath: 'workbook.xlsx')
→ Returns { batchId: 'abc123...' }

excel_powerquery(action: 'import', excelPath: 'workbook.xlsx', queryName: 'Sales', sourcePath: 'sales.pq', loadToWorksheet: false, batchId: 'abc123...')
excel_powerquery(action: 'set-load-to-data-model', excelPath: 'workbook.xlsx', queryName: 'Sales', batchId: 'abc123...')
excel_powerquery(action: 'import', excelPath: 'workbook.xlsx', queryName: 'Products', sourcePath: 'products.pq', loadToWorksheet: false, batchId: 'abc123...')
excel_powerquery(action: 'set-load-to-data-model', excelPath: 'workbook.xlsx', queryName: 'Products', batchId: 'abc123...')
... (2 more imports + loads)

commit_excel_batch(batchId: 'abc123...', save: true)

**SAVINGS: 75% faster (3-4 seconds vs 16-24 seconds)**

# Critical Rules

1. **DETECT PLURAL/NUMBERS IN REQUEST** → Use batch mode immediately
2. **Always call commit_excel_batch when done** (don't leave batches hanging)
3. **Use the same batchId for ALL operations on the same file**
4. **One batch per workbook** (can't share batches across different files)
5. **Set save: true to save changes, save: false to discard changes**

# Common Patterns

**Pattern 1: Multiple Power Query Imports**
User: 'Import sales.pq, products.pq, customers.pq, regions.pq'
→ 4 files = USE BATCH MODE

**Pattern 2: Create Multiple Parameters**
User: 'Create parameters for: Start_Date (2025-01-01), End_Date (2025-12-31), Region (West)'
→ 3 parameters = USE BATCH MODE

**Pattern 3: Multiple DAX Measures**
User: 'Create measures: TotalRevenue = SUM(Sales[Amount]), AvgPrice = AVERAGE(Sales[Price])'
→ 2 measures = USE BATCH MODE

# Response Hints (Warnings You Should NOT See)

If you see messages like:
- 'For multiple imports: Use begin_excel_batch to group operations efficiently'
- 'Creating multiple measures? Use begin_excel_batch to keep Data Model open'

**These mean you MISSED the batch mode opportunity.** Next time, detect the keywords earlier and use batch mode from the start.

# Decision Tree

START: Read user's request
↓
Count operations needed (look for numbers, plurals, lists)
↓
If 2+ operations on same file → BEGIN_EXCEL_BATCH first
If 1 operation only → Call tool directly (no batch)
↓
DONE: All operations complete → COMMIT_EXCEL_BATCH

REMEMBER: Batch mode is about UPFRONT DETECTION, not fixing after the fact!");
    }
}
