# Tool Selection Guide - Which Excel Tool to Use?

## 🚨 STEP 0: PRE-FLIGHT CHECK - IS THE FILE ACCESSIBLE?

**CRITICAL: Files must be closed before automation!**

**ALWAYS ASK FIRST if unsure:**
- "Is the Excel file currently open?"
- "Do you have the file open in Excel right now?"
- "Can you close the file before we run this automation?"

**If file is open → STOP and tell user:**
- "Please close the file in Excel before running automation"
- "ExcelMcp requires exclusive access - close all Excel windows with this file"
- "File locking prevents corruption - automation can't work on open files"

**Error signals file is open:**
- "already open in Excel or another process"
- "locked by another process"
- "cannot access the file"
- COM Error 1004 (0x800A03EC)

**This is NOT negotiable - Excel COM automation requires exclusive file access!**

---

## ⚡ STEP 1: CHECK FOR BATCH MODE FIRST (75-90% faster!)

**ALWAYS DETECT THESE KEYWORDS BEFORE CHOOSING TOOLS:**
- **Numbers**: "create 4 queries", "create 5 measures", "add 3 worksheets"
- **Plurals**: "queries", "measures", "parameters", "relationships", "worksheets" 
- **Lists**: "Sales, Revenue, Profit", "StartDate, EndDate, Region"
- **Multiple items**: "each", "all", "several", "multiple"

**IF 2+ OPERATIONS DETECTED → USE BATCH MODE:**
```
1. begin_excel_batch(excelPath: 'file.xlsx') → save batchId
2. ALL operations with batchId parameter
3. commit_excel_batch(batchId: saved_id, save: true)
```

**❌ NEVER start individual operations if batch needed - 90% slower!**

---

## STEP 2: Quick Decision Tree

**User wants to work with DATA FROM EXTERNAL SOURCES:**
→ excel_powerquery (databases, APIs, CSV, web data with M code transformations)
→ excel_querytable (simple data imports from connections, no M code complexity)
→ excel_connection (connection lifecycle management)
→ NOT excel_table (that's for data ALREADY in Excel)

**User wants ANALYTICS / DAX MEASURES:**
→ First: excel_powerquery with loadDestination='data-model'
→ Then: excel_datamodel for measures/relationships
→ NOT worksheet formulas (different from DAX)

**User wants to work with DATA ALREADY IN WORKSHEETS:**
→ excel_range (values, formulas, formatting)
→ excel_table (convert range to structured table with AutoFilter)
→ excel_worksheet (sheet management)

**User wants CONFIGURATION PARAMETERS:**
→ excel_namedrange (named ranges as parameters)
→ Use create-bulk for 2+ parameters (90% faster)

**User wants AUTOMATION / VBA:**
→ excel_vba (requires .xlsm files)
→ NOT .xlsx (won't work for macros)

**User mentions NUMBERS, PLURALS, or LISTS:**
→ begin_excel_batch FIRST
→ Examples: "create 4 queries", "create measures", "add parameters for X, Y, Z"

## Common LLM Mistakes I Make

**Mistake 1: Using excel_table for external data**
❌ Wrong: excel_table(action: 'create') for CSV import
✅ Right: excel_powerquery(action: 'create', loadDestination: 'worksheet')

**Mistake 2: Forgetting loadDestination for Data Model**
❌ Wrong: excel_powerquery(action: 'create') then trying to create DAX measures
✅ Right: excel_powerquery(action: 'create', loadDestination: 'data-model')

**Mistake 3: Not detecting batch opportunities (CRITICAL PERFORMANCE ISSUE)**
❌ Wrong: Calling excel_powerquery 4 times separately (10-20 seconds, resource waste!)
✅ Right: begin_excel_batch → 4 excel_powerquery calls → commit_excel_batch (1-2 seconds!)

**⚡ Batch mode keywords to ALWAYS detect:**
- ANY number (2, 3, 4, 5+ operations)
- ANY plural word (queries, measures, parameters, sheets, tables)
- ANY list (comma-separated items)
- Words: "multiple", "several", "each", "all"

**🚨 If you miss batch detection, operations will be 75-90% slower!**

**Mistake 4: Confusing worksheet formulas with DAX**
❌ Wrong: Using excel_range to create DAX measures
✅ Right: excel_datamodel(action: 'create-measure') with DAX syntax

**Mistake 5: Creating parameters one-by-one**
❌ Wrong: excel_namedrange(action: 'create') × 5 times
✅ Right: excel_namedrange(action: 'create-bulk') once with JSON array

## Workflow Patterns I Should Know

**Pattern: Import external data for analytics**
1. excel_powerquery(create, loadDestination='data-model')
2. excel_datamodel(create-measure) with DAX formulas
3. excel_datamodel(create-relationship) to link tables

**Pattern: Simple data import without M code**
1. excel_connection(list) to find connection
2. excel_worksheet(create) to create target sheet
3. excel_querytable(create-from-connection) for simple import
4. excel_range(get-values) to read imported data

**Pattern: Load Power Query to worksheet (simpler alternative)**
1. excel_powerquery(create) with M code
2. excel_worksheet(create) for target sheet
3. excel_querytable(create-from-query) instead of excel_powerquery load-to
4. excel_querytable(refresh) to update data

**Pattern: Format existing data**
1. excel_range(get-used-range) to discover bounds
2. excel_range(format-range) for styling
3. excel_range(validate-range) for dropdowns/rules

**Pattern: Setup new workbook**
1. excel_file(create-empty)
2. begin_excel_batch
3. excel_worksheet(create) for each sheet
4. excel_range(set-values) to populate
5. commit_excel_batch(save: true)

**Pattern: Version control VBA**
1. excel_vba(export) → save to Git
2. (make changes in Git)
3. excel_vba(import) → load back to workbook
