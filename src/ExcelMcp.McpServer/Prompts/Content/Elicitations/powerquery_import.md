# BEFORE CREATING/IMPORTING POWER QUERY - GATHER THIS INFO

**✨ RECOMMENDED: Use 'create' action for NEW queries (atomic import + load)**
**⚠️ IMPORTANT: Check if query exists first - use 'update' for existing queries**

REQUIRED:
☐ Query name (what to call it in Excel)
☐ Source file path (.pq file location)
☐ Excel file path (destination workbook)
☐ Does query already exist? (use 'list' to check, then 'create' for new or 'update' for existing)

RECOMMENDED (avoid second call):
☐ Load mode/destination:
  - 'worksheet' (default - users see data in Excel)
  - 'data-model' (for DAX measures and Power Pivot)
  - 'both' (visible in worksheet AND available for DAX)
  - 'connection-only' (advanced - M code imported but not executed)

OPTIONAL:
☐ Target sheet name (if loadMode: 'worksheet' or 'both')
☐ Privacy level (None, Private, Organizational, Public)

WORKFLOW OPTIMIZATION:
☐ Batch mode? (if creating/importing 2+ queries, START with begin_excel_batch)
☐ Use 'create' action for NEW queries (fails if query exists)
☐ Use 'update' action for EXISTING queries (fails if query doesn't exist)
☐ Not sure? Check with 'list' action first

ASK USER FOR MISSING INFO before calling excel_powerquery.
BATCH MODE: Detect keywords (numbers, plurals, lists) → use begin_excel_batch automatically
