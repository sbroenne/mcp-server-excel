using System.ComponentModel;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Server;

namespace Sbroenne.ExcelMcp.McpServer.Prompts;

/// <summary>
/// MCP prompt for helping LLMs choose the right Excel tool for the task.
/// </summary>
[McpServerPromptType]
public static class ExcelToolSelectionPrompts
{
    /// <summary>
    /// Comprehensive guide for selecting the appropriate Excel tool.
    /// </summary>
    [McpServerPrompt(Name = "excel_tool_selection_guide")]
    [Description("Guide for choosing the right Excel tool based on the user's request")]
    public static ChatMessage ToolSelectionGuide()
    {
        return new ChatMessage(ChatRole.User, @"# EXCEL TOOL SELECTION GUIDE

Choose the RIGHT tool for each Excel task. This guide helps you pick the most efficient tool.

## TOOL DECISION TREE

### DATA SOURCE QUESTIONS:

**Q: Where is the data coming from?**
- External source (database, web, file, API) → **excel_powerquery**
- Already in Excel worksheet (range of cells) → **excel_table** or **excel_range**
- User input/manual entry → **excel_parameter** or **excel_range**

**Q: Does it need to refresh from source?**
- Yes, periodic refresh needed → **excel_powerquery**
- No, static data → **excel_range** or **excel_table**

### OPERATION TYPE QUESTIONS:

**Q: What do you need to do?**

**Working with EXTERNAL DATA:**
→ **excel_powerquery** (M code transformations, data loading)
  - Import CSV, Excel files, databases, web APIs
  - Transform data with Power Query M code
  - Load to worksheet, Data Model, or both
  - Keywords: import, transform, load, refresh, M code, Power Query

**Working with DATA MODEL / ANALYTICS:**
→ **excel_datamodel** (DAX measures, relationships)
  - Create DAX measures for calculations
  - Build relationships between tables
  - Add calculated columns
  - Query Data Model structure
  - Keywords: DAX, measure, relationship, Power Pivot, Data Model, analytics

**Working with EXISTING DATA in WORKSHEET:**
→ **excel_range** (cell values, formulas)
  - Get/set cell values
  - Get/set formulas
  - Clear ranges
  - Copy/paste data
  - Find/replace
  - Keywords: cells, values, formulas, range, A1:Z100

→ **excel_table** (ListObject structure)
  - Convert range to structured table
  - Add AutoFilter
  - Sort and filter
  - Add/remove columns
  - Structured references ([@Column])
  - Keywords: table, ListObject, AutoFilter, sort, filter

**Working with CONFIGURATION:**
→ **excel_parameter** (named ranges for parameters)
  - Configuration values (dates, thresholds, names)
  - Reusable cell references
  - Bulk parameter creation
  - Keywords: parameter, configuration, named range, settings

**Working with CODE:**
→ **excel_vba** (Visual Basic macros)
  - Import/export VBA modules
  - Run macros
  - Version control for VBA
  - Keywords: VBA, macro, module, automate

**Working with WORKSHEETS:**
→ **excel_worksheet** (sheet lifecycle)
  - Create/delete sheets
  - Rename/copy sheets
  - List all sheets
  - Keywords: worksheet, sheet, tab

**Working with FILES:**
→ **excel_file** (file creation)
  - Create new blank workbooks
  - Keywords: create file, new workbook

## COMMON SCENARIOS

**Scenario 1: Import CSV and analyze with DAX**
1. excel_powerquery: import CSV → loadDestination: 'data-model'
2. excel_datamodel: create DAX measures

**Scenario 2: Create parameter-driven report**
1. excel_parameter: create-bulk → configuration values
2. excel_range: set formulas using parameter references

**Scenario 3: Version control VBA code**
1. excel_vba: export → save to Git
2. excel_vba: import → load from Git

**Scenario 4: Structure existing data**
1. excel_table: create → convert range to table
2. excel_table: apply-filter → add filters
3. excel_table: sort → organize data

**Scenario 5: Multi-query Data Model workflow**
1. begin_excel_batch → start session
2. excel_powerquery: import × 4 → loadDestination: 'data-model'
3. excel_datamodel: create-relationship × 3 → build model
4. excel_datamodel: create-measure × 5 → add calculations
5. commit_excel_batch → save all

## ANTI-PATTERNS (DON'T DO THIS)

❌ Using excel_table for external data
  → Use excel_powerquery instead

❌ Using excel_powerquery for data already in Excel
  → Use excel_table or excel_range instead

❌ Creating parameters one-by-one
  → Use excel_parameter create-bulk instead

❌ Calling tools without batch mode for 2+ operations
  → Use begin_excel_batch first

❌ Using set-load-to-table then trying excel_table add-to-datamodel
  → Use loadDestination: 'data-model' or 'both' on import

## TOOL FEATURE MATRIX

| Feature | excel_powerquery | excel_table | excel_range |
|---------|------------------|-------------|-------------|
| External data | ✅ | ❌ | ❌ |
| Refresh from source | ✅ | ❌ | ❌ |
| M code transformations | ✅ | ❌ | ❌ |
| AutoFilter | ❌ | ✅ | ❌ |
| Structured refs | ❌ | ✅ | ❌ |
| Get/set values | ❌ | ❌ | ✅ |
| Formulas | ❌ | ❌ | ✅ |
| Load to Data Model | ✅ | ✅ | ❌ |

REMEMBER: Always choose the tool that matches the data source and operation type!");
    }
}
