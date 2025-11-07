# BEFORE CREATING/IMPORTING POWER QUERY - GATHER THIS INFO

**✨ RECOMMENDED: Use 'create' action for new queries (atomic import + load)**

REQUIRED:
☐ Query name (what to call it in Excel)
☐ Source file path (.pq file location)
☐ Excel file path (destination workbook)

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
☐ Use 'create' action instead of 'import' for atomic operation
☐ Use 'update-and-refresh' instead of 'update' + 'refresh' for production updates

ASK USER FOR MISSING INFO before calling excel_powerquery.
BATCH MODE: Detect keywords (numbers, plurals, lists) → use begin_excel_batch automatically
