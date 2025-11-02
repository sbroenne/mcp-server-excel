# Tool Selection Guide - Which Excel Tool to Use?

## Quick Decision Tree

**User wants to work with DATA FROM EXTERNAL SOURCES:**
→ excel_powerquery (databases, APIs, CSV, web data)
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
→ Examples: "import 4 queries", "create measures", "add parameters for X, Y, Z"

## Common LLM Mistakes I Make

**Mistake 1: Using excel_table for external data**
❌ Wrong: excel_table(action: 'create') for CSV import
✅ Right: excel_powerquery(action: 'import', loadDestination: 'worksheet')

**Mistake 2: Forgetting loadDestination for Data Model**
❌ Wrong: excel_powerquery(action: 'import') then trying to create DAX measures
✅ Right: excel_powerquery(action: 'import', loadDestination: 'data-model')

**Mistake 3: Not detecting batch opportunities**
❌ Wrong: Calling excel_powerquery 4 times separately (slow!)
✅ Right: begin_excel_batch → 4 excel_powerquery calls → commit_excel_batch

**Mistake 4: Confusing worksheet formulas with DAX**
❌ Wrong: Using excel_range to create DAX measures
✅ Right: excel_datamodel(action: 'create-measure') with DAX syntax

**Mistake 5: Creating parameters one-by-one**
❌ Wrong: excel_namedrange(action: 'create') × 5 times
✅ Right: excel_namedrange(action: 'create-bulk') once with JSON array

## Workflow Patterns I Should Know

**Pattern: Import external data for analytics**
1. excel_powerquery(import, loadDestination='data-model')
2. excel_datamodel(create-measure) with DAX formulas
3. excel_datamodel(create-relationship) to link tables

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
