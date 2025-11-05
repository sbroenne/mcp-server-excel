# BEFORE CREATING QUERYTABLE - GATHER THIS INFO

REQUIRED:
☐ Worksheet name (where to create QueryTable)
☐ QueryTable name (what to call it)
☐ Source type:
  - 'connection' → need connection name (use excel_connection list)
  - 'query' → need Power Query name (use excel_powerquery list)

FOR CREATE-FROM-CONNECTION:
☐ Connection name (from excel_connection list)
☐ Sheet must exist (use excel_worksheet create if needed)

FOR CREATE-FROM-QUERY:
☐ Power Query name (from excel_powerquery list)
☐ Sheet must exist (use excel_worksheet create if needed)

RECOMMENDED (avoid second call):
☐ Refresh settings:
  - refreshImmediately (default: true - immediate data load)
  - refreshOnFileOpen (default: false)
  - backgroundQuery (default: false)
☐ Formatting options:
  - preserveColumnInfo (default: true)
  - preserveFormatting (default: true)
  - adjustColumnWidth (default: true)

OPTIONAL:
☐ Range address (default: 'A1')
☐ Save password setting (default: false - recommended for security)

WORKFLOW OPTIMIZATION:
☐ Batch mode? (if creating 2+ QueryTables, START with begin_excel_batch)
☐ Sheet exists? (use excel_worksheet list to check, create if needed)
☐ Connection/Query exists? (use excel_connection/excel_powerquery list to verify)

COMMON PATTERN:
1. excel_worksheet list → check if sheet exists
2. excel_worksheet create (if needed)
3. excel_connection list OR excel_powerquery list
4. excel_querytable create-from-connection/create-from-query
5. excel_range get-values to read imported data

ASK USER FOR MISSING INFO before calling excel_querytable.
BATCH MODE: Detect keywords (numbers, plurals, lists) → use begin_excel_batch automatically
